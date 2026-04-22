import fitz
import re
import io
import json
import time
from pathlib import Path
from PIL import Image
import pytesseract
from googleapiclient.discovery import build
from google.oauth2 import service_account
from googleapiclient.http import MediaIoBaseDownload


FOLDER_ID = "1J22Hv9BJD5AoB-jCepMMQgEriM-eIVnq"
SCOPES = ['https://www.googleapis.com/auth/drive.readonly']
CREDS_PATH = "credentials.json"  # Use relative path - file is in same directory
OUTPUT_FILE = "structured_results.json"
REFERENCE_FILE = "vehicle_reference.json"


OIL_PATTERN = r"\b(0|5|10|15|20|25)\s*W[-–—]?\s*(16|20|30|40|50|60)\b"
CAPACITY_PATTERN = r"(\d+\.?\d*)\s*(?:us\s+|imp\s+|u\.s\.\s+)?(quarts?|qts?|qt\.?|liters?|l\b)"
CAPACITY_UNIT_PATTERN = r"(?:quarts?|qts?|qt\.?|liters?|litres?|l\b|gal|gallons?|ml|cc)"
# Match engine sizes: "2.5L", "2.5 liter", "2.5 GDI", "2.5 DOHC", "2.5-cylinder", etc.
ENGINE_PATTERN = r"\b([1-6]\.\d)\s*(?:[-]?\s*(?:l|liter|litre)|(?:\s+(?:gdi|dohc|sohc|turbo|ecoboost|naturally|cylinder)))\b"
# Generic engine-type detector without a static token list.
ENGINE_TYPE_PATTERN = r"\b(?:[viwhf]\s*-?\s*\d{1,2}|inline\s*-?\s*[3-8]|flat\s*-?\s*(?:3|4|6|8|12)|boxer|rotary|wankel|turbo(?:charged)?|supercharged|naturally\s*-?\s*aspirated|hybrid|electric|diesel|petrol|gdi|sohc|dohc)\b"
TEMP_PATTERN = r"(-?\d+)\s*°?\s*(c|f)"  # Matches temps like -30C, 100°F, 50 F, etc.

NON_ENGINE_CONTEXT = [
    "brake", "transmission", "gear oil", "power steering",
    "differential", "coolant", "washer", "clutch", "mtf",
    "manual transmission fluid", "temporary replacement", "filler bolt",
    "synchronizer", "gearbox fluid", "transmission fluid", "atf",
    "automatic transmission", "fluid level", "fluid change",
    "fuel", "gallon", "tank", "capacity"
]

INVALID_WORDS = [
    "motor", "motors", "company", "co", "ltd", "inc", "manual", "owner"
]

OIL_PAGE_KEYWORDS = [
    "oil capacity", "engine oil", "crankcase", "oil with filter",
    "oil change capacity", "including filter", "engine oil capacity",
    "engine oil recommendation", "viscosity", "api service",
    "lubricant", "specifications", "technical information"
]


def has_engine_signal(text):
    """
    Dynamic engine-context detector for lines/sentences.
    Uses displacement pattern, comprehensive engine type pattern, and core engine terms.
    """
    if not text:
        return False

    text_lower = text.lower()

    # Primary signals: actual engine sizes and engine type tokens.
    if re.search(ENGINE_PATTERN, text_lower):
        return True
    if re.search(ENGINE_TYPE_PATTERN, text_lower):
        return True

    # Secondary context terms for table-style specs.
    return bool(re.search(r"\b(engine|displacement|cylinder|liter|litre|cc|crankcase|horsepower|hp|ecoboost|vortec|duratec|duratech|modular|triton)\b", text_lower))


def is_capacity_conversion_engine_match(text, match):
    """
    Return True when an engine regex match is really a metric capacity conversion.
    Example: in "Engine Oil with Filter 4.0 qt 3.8 L", "3.8 L" is a capacity,
    not a 3.8L engine.
    """
    if not text or not match:
        return False

    context_before = text[max(0, match.start() - 40):match.start()]
    return bool(
        re.search(r"\d+\.?\d*\s*" + CAPACITY_UNIT_PATTERN + r"\s*\(?\s*$", context_before, re.I)
        or re.search(r"\(\s*$", context_before, re.I)
    )


def is_parenthesized_capacity_conversion(text, match):
    """
    True for metric conversions like "6.0 quarts (5.7L)".
    """
    if not text or not match:
        return False

    before = text[max(0, match.start() - 35):match.start()]
    after = text[match.end():min(len(text), match.end() + 5)]
    return bool(
        re.search(r"\d+\.?\d*\s*(?:quarts?|qts?|qt\.?|gal|gallons?)\s*\(\s*$", before, re.I)
        and re.search(r"^\s*\)", after)
    )


def is_capacity_or_fluid_row(text):
    """
    Identify capacity/fluid table rows so they are not mistaken for engine spec rows.
    This stays generic: it checks fluid/capacity vocabulary, not vehicle-specific engines.
    """
    if not text:
        return False

    text_lower = text.lower()
    has_non_l_capacity_unit = bool(re.search(
        r"\d+\.?\d*\s*(?:quarts?|qts?|qt\.?|gal|gallons?|ml|cc)",
        text_lower,
        re.I
    ))
    has_metric_conversion = bool(re.search(
        r"\d+\.?\d*\s*(?:quarts?|qts?|qt\.?|gal|gallons?)\s+\d+\.?\d*\s*(?:liters?|litres?|l\b)",
        text_lower,
        re.I
    ))
    has_fluid_term = any(term in text_lower for term in [
        "engine oil", "oil with filter", "cooling system", "coolant",
        "fuel tank", "transaxle fluid", "transmission fluid",
        "differential", "power steering", "refrigerant", "washer fluid"
    ])

    return has_fluid_term or has_non_l_capacity_unit or has_metric_conversion


def is_real_capacity_match(text, match):
    """
    Distinguish real fluid capacities from engine displacements that also look like liters.
    A bare "6.2L engine" is not a capacity; "8.0 L 8.5 qt" or "4.0 Quarts" is.
    """
    if not text or not match:
        return False

    unit = match.group(2).lower()
    if 'qt' in unit or 'quart' in unit or 'gal' in unit or unit in {'ml', 'cc'}:
        return True

    before = text[max(0, match.start() - 30):match.start()]
    after = text[match.end():min(len(text), match.end() + 30)]

    paired_with_quarts = (
        re.search(r"\d+\.?\d*\s*(?:quarts?|qts?|qt\.?)\s*$", before, re.I)
        or re.search(r"^\s*\d+\.?\d*\s*(?:quarts?|qts?|qt\.?)", after, re.I)
    )
    if paired_with_quarts:
        return True

    if re.search(r"^\s*(?:engine|engines|v\d|i\d|turbo|flex fuel)", after, re.I):
        return False

    if unit.startswith('l'):
        return False

    return any(term in text.lower() for term in [
        "capacity", "capacities", "with filter", "fluid", "cooling system",
        "fuel tank", "quarts", "qts", "qt"
    ])

def get_temperature_with_fallback(extracted_temps, oil_type):
    """
    Get temperature condition for an oil - DYNAMIC from text with smart fallback.
    
     Logic:
     1. Use temperatures extracted directly from the PDF text
     2. Normalize casing/spacing for consistency
     3. If nothing is extracted, return a generic dynamic-safe fallback
    
    Args:
        extracted_temps: Set of temps extracted from PDF text
        oil_type: Oil type like "0W-20", "5W-30", etc.
    
    Returns:
        Set of temperature conditions to display
    """
    if extracted_temps:
        normalized = {
            str(t).strip().lower()
            for t in extracted_temps
            if str(t).strip()
        }
        if normalized:
            return normalized

    # Dynamic fallback when no explicit temperature text is found.
    return {"all temperatures"}


def has_non_engine_oil_context(text):
    """
    Decide whether an oil mention belongs to another fluid system.

    This stays dynamic and phrase-based. It avoids broad substring checks like
    "fuel", which can appear in valid engine-oil prose such as "fuel economy".
    """
    if not text:
        return False

    text_lower = text.lower()

    non_engine_phrases = [
        "manual transmission fluid",
        "automatic transmission fluid",
        "transmission fluid",
        "power steering fluid",
        "brake fluid",
        "gear oil",
        "coolant",
        "washer fluid",
        "differential",
        "temporary replacement",
        "dexron",
        "atf",
        "fuel tank capacity",
    ]
    return any(phrase in text_lower for phrase in non_engine_phrases)


def has_engine_oil_context(text):
    """
    Require positive evidence that an oil mention belongs to engine-oil guidance.

    This stays document-driven and generic: it looks for common engine-oil
    phrasing rather than any specific make/model data.
    """
    if not text:
        return False

    text_lower = text.lower()
    if has_non_engine_oil_context(text_lower):
        return False

    engine_oil_signals = [
        "engine oil",
        "recommended engine oil",
        "api certification seal",
        "for gasoline engines",
        "viscosity or weight",
        "select the oil for your car",
        "oil with a viscosity of",
        "oil viscosity",
        "year-round protection",
        "improved fuel economy",
        "oil container",
    ]
    return any(signal in text_lower for signal in engine_oil_signals)



def extract_text_from_images(doc, pages_with_images):
    """
    Extract text from PDF pages containing images using OCR via pytesseract.
    
    Strategy:
    1. Iterate through pages with images
    2. Render each page to an image at high resolution (300 DPI)
    3. Use pytesseract to extract text from the image
    4. Combine all extracted text with proper spacing
    
    Args:
        doc: PyMuPDF document object
        pages_with_images: List of page numbers containing images
    
    Returns:
        String containing all OCR'd text from image pages
    """
    ocr_text = []
    
    for page_num in pages_with_images:
        try:
            page = doc[page_num]
            # Render page to image at high resolution (300 DPI = 4x color resolution)
            pix = page.get_pixmap(matrix=fitz.Matrix(4, 4))
            # Convert pixmap to PIL Image
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            # Extract text from image using pytesseract
            text = pytesseract.image_to_string(img)
            if text.strip():
                ocr_text.append(text)
        except Exception as e:
            print(f"      Warning: OCR failed on page {page_num}: {str(e)}")
            continue
    
    return " ".join(ocr_text)


def get_drive_service():
    """
    Step 1: Authenticate with Google Drive
    Loads service account credentials and returns an authenticated Drive API client.
    This allows reading files from Google Drive folders.
    """
    creds = service_account.Credentials.from_service_account_file(
        CREDS_PATH, scopes=SCOPES
    )
    return build("drive", "v3", credentials=creds)


def get_all_pdfs(service, folder_id):
    """
    Step 2: Recursively search Google Drive folder for PDF files
    Traverses all subfolders and collects PDF files.
    
    Args:
        service: Authenticated Google Drive API client
        folder_id: Root folder ID to search
    
    Returns:
        List of PDF file objects with id, name, mimeType
    """
    pdfs, folders = [], [folder_id]
    
    while folders:
        current = folders.pop()
        response = service.files().list(
            q=f"'{current}' in parents and trashed=false",
            fields="files(id,name,mimeType)",
            supportsAllDrives=True,
            includeItemsFromAllDrives=True
        ).execute()
        
        for item in response.get("files", []):
            if item["mimeType"] == "application/vnd.google-apps.folder":
                folders.append(item["id"])
            elif item["mimeType"] == "application/pdf":
                pdfs.append(item)
    
    return pdfs


def analyze_pdf_type(doc):
    """
    Determine if PDF is text-heavy (use text extraction) or manual/spec (use OCR).
    
    Strategy:
    1. Calculate average text characters per page
    2. If > 800 chars/page: Text-heavy (AUTO extraction)
    3. If < 800 chars/page: Manual/spec document (use OCR)
    
    Args:
        doc: PyMuPDF document object
    
    Returns:
        Tuple: (extraction_type, avg_chars_per_page)
            extraction_type: "AUTO" for text extraction, "MANUAL" for OCR
            avg_chars_per_page: Average characters per page as integer
    """
    total_chars = 0
    page_count = len(doc)
    
    for page in doc:
        text = page.get_text("text")
        total_chars += len(text)
    
    avg_chars = int(total_chars / page_count) if page_count > 0 else 0
    
    # Text-heavy PDFs have lots of direct text extraction potential
    # Manual/spec documents are mostly images and need OCR
    extraction_type = "AUTO" if avg_chars >= 800 else "MANUAL"
    
    return extraction_type, avg_chars


def download_pdf(service, file_id):
    """
    Step 3: Download PDF file from Google Drive
    Retrieves file content as binary stream.
    
    Args:
        service: Authenticated Google Drive API client
        file_id: Google Drive file ID
    
    Returns:
        BytesIO buffer containing PDF data
    """
    request = service.files().get_media(fileId=file_id)
    buffer = io.BytesIO()
    downloader = MediaIoBaseDownload(buffer, request)
    done = False
    
    while not done:
        _, done = downloader.next_chunk()
    
    buffer.seek(0)
    return buffer


def clean_text(text):
    """
    Clean PDF text by normalizing whitespace and line breaks.
    Converts multiple spaces/newlines to single spaces for uniform parsing.
    
    Args:
        text: Raw PDF extracted text
    
    Returns:
        Cleaned text string
    """
    return re.sub(r"\s+", " ", text.replace("\n", " "))


def normalize_oil(raw):
    """
    Standardize oil type format (e.g., "5w-30" becomes "5W-30").
    Handles different dash styles and ensures consistent spacing.
    
    Args:
        raw: Raw oil type string from PDF
    
    Returns:
        Normalized oil type string
    """
    oil = raw.upper().replace("–", "-").replace("—", "-")
    oil = re.sub(r"\s+", "", oil)
    if "W-" not in oil:
        oil = oil.replace("W", "W-")
    return oil


def to_quarts_liters(value, unit):
    """
    Convert capacity value to both quarts and liters.
    Handles imperial (quarts) to metric (liters) conversion and vice versa.
    
    Args:
        value: Numeric capacity value
        unit: Unit type (quarts, qts, liters, l, etc.)
    
    Returns:
        Tuple of (quarts, liters)
    """
    value = float(value)
    if unit.lower().startswith("l"):
        return round(value / 0.946, 2), value
    else:
        return value, round(value * 0.946, 1)



def parse_filename(name):
    """
    Extract vehicle info from PDF filename format.
    Handles filenames with optional "Copy of " prefix.
    Expected format: [Copy of ]YYYY-Make-Model.pdf (e.g., 2017-Honda-Civic.pdf)
                    or: Copy of 2017-Honda-Civic.pdf
    
    Args:
        name: Filename string
    
    Returns:
        Tuple of (year, make, model) or (None, None, None) if parsing fails
    """
    # Remove "Copy of " prefix if present
    clean_name = name.replace("Copy of ", "").strip()
    
    # Match YYYY-Make-Model-Suffix.pdf pattern
    match = re.match(r"(\d{4})-([^-]+)-(.+)\.pdf", clean_name, re.I)
    if match:
        year, make, model = match.groups()
        # Remove common suffixes like -OM (Owner's Manual), -UG (User Guide)
        model = model.replace("-OM", "").replace("-UG", "").replace("-UM", "")
        return int(year), make.capitalize(), model.capitalize()
    return None, None, None


def build_multi_engine_data(engine_caps, oil_scores, oil_temps, engine_oil_map):
    """
    Build structured oil recommendations for each engine in vehicle.
    Calculates per-engine primary/secondary oil selection based on scoring.
    
    Handles engine variants like "1.4L", "1.4L Turbo", "5.3L V8", etc.
    
    Step-by-step process:
    1. Iterate through each detected engine size
    2. Get valid oils from proximity mapping or size-based assignment
    3. Calculate best-scoring oil for THIS engine (not globally)
    4. Mark high-scoring oils as primary/secondary
    5. Include oils with actual temperature conditions
    6. Build complete engine data structure with capacities
    
    Args:
        engine_caps: Dict of engine sizes to capacity data
        oil_scores: Dict of oils to recommendation scores
        oil_temps: Dict of oils to temperature sets
        engine_oil_map: Dict of engines to nearby oils
    
    Returns:
        Dict with engine-to-recommendations structure
    """
    engine_data = {}

    for eng, cap in engine_caps.items():
        with_filter = cap.get("with_filter")
        without_filter = cap.get("without_filter")
        oil_list = []

        if oil_scores:
            # First try direct engine lookup
            valid_oils = engine_oil_map.get(eng, [])
            
            # If not found, try matching base engine size (e.g., '6.0L' from '6.0L V8')
            if not valid_oils:
                base_eng_size = eng.split()[0]  # Get "6.0L" from "6.0L V8"
                valid_oils = engine_oil_map.get(base_eng_size, [])
            
            # If still not found, extract numeric engine size for fallback filtering
            if not valid_oils:
                # Extract numeric engine size from variants like "1.4L Turbo"
                eng_size = 0
                if eng != "unknown_engine":
                    match = re.search(r'(\d+\.?\d*)', eng)
                    if match:
                        try:
                            eng_size = float(match.group(1))
                        except (ValueError, AttributeError):
                            eng_size = 0
                
                if eng_size > 0:
                    for oil in oil_scores.keys():
                        if eng_size <= 1.6 and "40" in oil:
                            valid_oils.append(oil)
                        elif eng_size >= 2.0 and "20" in oil:
                            valid_oils.append(oil)
                else:
                    valid_oils = list(oil_scores.keys())

            if valid_oils:
                engine_best = max(valid_oils, key=lambda oil: oil_scores.get(oil, 0))
                engine_max_score = oil_scores.get(engine_best, 0)
            else:
                engine_best = max(oil_scores, key=oil_scores.get)
                engine_max_score = oil_scores[engine_best]
                valid_oils = list(oil_scores.keys())

            for oil in valid_oils:
                score = oil_scores.get(oil, 0)
                temps = oil_temps.get(oil, [])
                
                has_actual_temps = any(
                    ("°F" in t or "°C" in t or "F" in t or "C" in t or 
                     "above" in t or "below" in t or "range:" in t or 
                     "weather" in t or "temperatures" in t or "cold" in t or "hot" in t)
                    for t in temps
                )

                if score >= engine_max_score - 1 or has_actual_temps:
                    oil_list.append({
                        "oil_type": oil,
                        "recommendation_level": "primary" if oil == engine_best else "secondary",
                        "temperature_condition": list(temps),
                    })

        engine_data[eng] = {
            "oil_capacity": {
                "with_filter": {k: v for k, v in with_filter.items() if k != "pos"} if with_filter else None,
                "without_filter": {k: v for k, v in without_filter.items() if k != "pos"} if without_filter else None
            },
            "oil_recommendations": oil_list
        }

    return engine_data


def expand_shared_capacity_to_detected_engines(engine_caps, all_engines):
    """
    Some manuals list one shared "Engine Oil with Filter" capacity, then list the
    engine options elsewhere. In that case the capacity is extracted as
    unknown_engine first, then dynamically copied to the detected engine list.
    """
    if not engine_caps or "unknown_engine" not in engine_caps or not all_engines:
        return engine_caps

    shared_cap = engine_caps.pop("unknown_engine")
    existing_bases = {eng.split()[0] for eng in engine_caps.keys()}

    for eng in all_engines:
        eng_base = eng.split()[0]
        if eng_base in existing_bases:
            continue

        engine_caps[eng] = {
            "with_filter": dict(shared_cap["with_filter"]) if shared_cap.get("with_filter") else None,
            "without_filter": dict(shared_cap["without_filter"]) if shared_cap.get("without_filter") else None
        }
        existing_bases.add(eng_base)

    return engine_caps


def align_capacity_engine_keys_with_detected_variants(engine_caps, all_engines):
    """
    If capacities are keyed by base displacement but engine detection found a typed
    variant for the same base, use the richer detected key.
    """
    if not engine_caps or not all_engines:
        return engine_caps

    variant_by_base = {}
    for eng in all_engines:
        base = eng.split()[0]
        if len(eng.split()) > 1:
            variant_by_base[base] = eng

    aligned = {}
    for eng, cap in engine_caps.items():
        if eng == "unknown_engine":
            aligned[eng] = cap
            continue

        base = eng.split()[0]
        aligned_key = variant_by_base.get(base, eng)
        aligned[aligned_key] = cap

    return aligned


def pair_quarts_liters(cap_list):
    """
    Pair capacity values by matching quarts with corresponding liters.
    Assumes alternating quarts/liters entries in list.
    
    Step-by-step process:
    1. Iterate through capacity list in pairs
    2. Match quarts value with following liters value
    3. Skip unpaired entries
    4. Return list of {"quarts": X, "liters": Y} objects
    
    Args:
        cap_list: List of capacity objects with quarts/liters fields
    
    Returns:
        List of paired capacity objects
    """
    paired = []
    i = 0
    
    while i < len(cap_list) - 1:
        q = cap_list[i]
        l = cap_list[i + 1]

        if q["quarts"] and l["liters"]:
            paired.append({
                "quarts": q["quarts"],
                "liters": l["liters"]
            })
            i += 2
        else:
            i += 1

    return paired


_vehicle_reference_cache = None


def load_vehicle_reference():
    """
    Load optional vehicle make/model reference data from disk.
    Kept separate from parsing logic so it remains a fallback only.
    """
    global _vehicle_reference_cache
    if _vehicle_reference_cache is not None:
        return _vehicle_reference_cache

    ref_path = Path(__file__).with_name(REFERENCE_FILE)
    if not ref_path.exists():
        _vehicle_reference_cache = {}
        return _vehicle_reference_cache

    try:
        with ref_path.open("r", encoding="utf-8") as f:
            data = json.load(f)
    except (OSError, json.JSONDecodeError):
        data = {}

    normalized = {}
    if isinstance(data, dict):
        for make, models in data.items():
            if not isinstance(make, str) or not isinstance(models, list):
                continue
            clean_models = [m for m in models if isinstance(m, str) and m.strip()]
            if clean_models:
                normalized[make.strip()] = clean_models

    _vehicle_reference_cache = normalized
    return _vehicle_reference_cache


def normalize_vehicle_label(value):
    """
    Normalize make/model text for reference matching.
    """
    if not value:
        return ""

    text = value.lower()
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def simplify_model_for_lookup(model):
    """
    Reduce detected model text to a stable lookup label.
    Example: "Civic Hatchback" -> "civic"
    """
    normalized = normalize_vehicle_label(model)
    if not normalized:
        return ""

    body_styles = {
        "hatchback", "sedan", "coupe", "wagon", "suv", "truck", "van",
        "pickup", "convertible", "roadster", "cabriolet", "fastback"
    }
    tokens = [token for token in normalized.split() if token not in body_styles]
    if not tokens:
        return normalized
    return " ".join(tokens)


def is_make_suspicious(make):
    """
    Identify low-confidence make values that should be eligible for fallback repair.
    """
    normalized = normalize_vehicle_label(make)
    if not normalized:
        return True

    suspicious_tokens = set(INVALID_WORDS) | {
        "and", "you", "your", "the", "this", "that", "with", "from",
        "for", "use", "using", "service", "manual", "owner", "owners",
        "guide", "information", "summary", "engine", "gasoline"
    }
    return normalized in suspicious_tokens


def is_known_make(make, reference=None):
    """
    Check whether a detected make matches a known make in the optional reference file.
    """
    if not make:
        return False

    if reference is None:
        reference = load_vehicle_reference()

    make_key = normalize_vehicle_label(make)
    known_keys = {normalize_vehicle_label(ref_make) for ref_make in reference.keys()}
    return make_key in known_keys


def resolve_make_from_model_reference(model, detected_make=None):
    """
    Resolve make from model using the optional reference file.
    Used only as a fallback when the detected make is missing or suspicious.
    """
    if not model:
        return detected_make

    reference = load_vehicle_reference()
    if not reference:
        return detected_make

    if detected_make and not is_make_suspicious(detected_make) and is_known_make(detected_make, reference):
        return detected_make

    model_key = simplify_model_for_lookup(model)
    if not model_key:
        return detected_make

    candidates = []
    for make, models in reference.items():
        for ref_model in models:
            ref_key = normalize_vehicle_label(ref_model)
            if not ref_key:
                continue

            if model_key == ref_key or model_key.startswith(ref_key + " ") or ref_key.startswith(model_key + " "):
                candidates.append(make)
                break

    unique_candidates = sorted(set(candidates))
    if len(unique_candidates) == 1:
        return unique_candidates[0]

    return detected_make


def detect_vehicle_from_pdf(doc):
    """
    Extract vehicle year, make, model from first few PDF pages.
    Falls back to filename parsing when PDF extraction fails.
    
    Step-by-step process:
    1. Extract text from first 5 pages
    2. Search for 4-digit year (1900-2099)
    3. Rank frequent candidate tokens from the document text
    4. Find body type (hatchback, sedan, etc.)
    5. Build model name near body type
    6. Fallback to word pair detection for make/model
    
    Args:
        doc: PyMuPDF document object
    
    Returns:
        Tuple of (year, make, model)
    """
    text = ""
    for i in range(min(5, len(doc))):
        text += doc[i].get_text()
    
    text_lower = text.lower()
    words = re.findall(r"[a-z]{3,}", text_lower)

    year_match = re.search(r"(19|20)\d{2}", text_lower)
    year = int(year_match.group()) if year_match else None

    # Dynamic token ranking from first pages (no fixed make/model dictionary).
    stop_words = set(INVALID_WORDS) | {
        "manual", "owners", "owner", "guide", "service", "maintenance",
        "vehicle", "vehicles", "specification", "specifications", "capacity", "oil",
        "and", "for", "the", "with", "your", "you", "this", "that", "from", "page",
        "pages", "use", "using", "recommended", "information", "summary",
        "gasoline", "engine", "api", "seal", "may", "never", "goes", "below",
        "above", "temperature", "temperatures", "protection", "improved", "economy"
    }
    token_freq = {}
    for w in words:
        if w in stop_words or len(w) < 3:
            continue
        token_freq[w] = token_freq.get(w, 0) + 1

    ranked_tokens = [w for w, _ in sorted(token_freq.items(), key=lambda x: x[1], reverse=True)]
    make = None
    for token in ranked_tokens:
        if token not in stop_words:
            make = token.capitalize()
            break

    model = None
    make_model_body_match = re.search(
        r"\b([a-z]{3,})\s+([a-z]{3,})\s+(hatchback|sedan|coupe|wagon|suv|truck|van|pickup)\b",
        text_lower
    )
    if make_model_body_match:
        make_candidate = make_model_body_match.group(1)
        model_token = make_model_body_match.group(2)
        body_type = make_model_body_match.group(3)
        if make_candidate not in stop_words:
            make = make_candidate.capitalize()
        model = f"{model_token.capitalize()} {body_type.capitalize()}"
    else:
        body_match = re.search(r"\b([a-z]{3,})\s+(hatchback|sedan|coupe|wagon|suv|truck|van|pickup)\b", text_lower)
        if body_match:
            model = f"{body_match.group(1).capitalize()} {body_match.group(2).capitalize()}"
        else:
            for token in ranked_tokens:
                if token not in stop_words and (make is None or token.capitalize() != make):
                    model = token.capitalize()
                    break

    make = resolve_make_from_model_reference(model, make)
    
    return year, make, model


def map_oils_to_engines(text):
    """
    Associate oils with nearby engines based on text proximity.
    Searches 200 chars before and 300 chars after each engine mention.
    
    Also attempts to match engine variants properly (e.g., "1.4L Turbo" → "1.4L").
    
    Step-by-step process:
    1. Find all engine size mentions in text with positions
    2. For each engine, extract text window around it
    3. Find oils in that window
    4. Store list of oils per engine
    5. Try to match variants to base engine sizes
    
    Args:
        text: Full PDF text
    
    Returns:
        Dict mapping engine sizes to lists of nearby oils
    """
    engine_oil_map = {}
    engine_positions = []
    
    for m in re.finditer(ENGINE_PATTERN, text):
        eng = f"{float(m.group(1)):.1f}L"
        engine_positions.append((eng, m.start()))

    for eng, eng_pos in engine_positions:
        window = text[max(0, eng_pos - 200): eng_pos + 300]
        oils = re.findall(OIL_PATTERN, window)
        engine_oil_map[eng] = list(set([
            normalize_oil(f"{b}W-{g}") for b, g in oils
        ]))

    return engine_oil_map


def extract_engines(text):
    """
    Extract engine sizes from PDF text with engine specification context.
    Also extracts engine variants (Turbo, V8, supercharged, etc.)
    
    DYNAMIC APPROACH:  Extracts engines from engine-related sentences, 
    which are then filtered by capacity matching. No hardcoding.
    
    Step-by-step process:
    1. Split text into sentences for context
    2. Filter sentences with engine keywords
    3. Skip sentences about non-specs (towing, cargo, previous gen, etc)
    4. Extract engine sizes and variants from relevant sentences
    5. CRITICAL: Skip engine matches that follow capacity indicators (qt, liter, gal, etc)
    6. Validate displacement is 0.8 to 8.0 liters
    7. Return list of found engines (will be filtered by capacities)
    
    Args:
        text: Full PDF text
    
    Returns:
        List of unique engine sizes (e.g., ["3.4L", "3.4L V6", "5.3L V8"])
    """
    engines = []
    text_lower = text.lower()
    
    skip_phrases = [
        "previous generation", "previous model", "prior generation",  
        "towing capacity", "cargo capacity", "payload capacity",
        "weight rating", "gvwr", "gcwr", "optional", "as an option",
        "engine oil", "transaxle fluid", "cooling system", "fuel tank"  # ADDED: Skip capacity sections
    ]
    
    # Split into sentences
    sentences = re.split(r'[.!?]\s+', text)
    
    for sentence in sentences:
        sentence_lower = sentence.lower()
        
        # Skip if mentions non-current or non-engine specs
        if any(phrase in sentence_lower for phrase in skip_phrases):
            continue
        
        # Must contain dynamic engine signal
        if not has_engine_signal(sentence_lower):
            continue
        
        # Skip non-engine context (fuel, tanks, etc)
        if any(ctx in sentence_lower for ctx in NON_ENGINE_CONTEXT):
            continue
        
        # Extract engines from this sentence
        for m in re.finditer(ENGINE_PATTERN, sentence_lower):
            try:
                num = float(m.group(1).strip())
                if 0.8 <= num <= 8.0:
                    eng_str = f"{num:.1f}L"
                    
                    # CRITICAL CHECK: Skip if this engine match appears AFTER capacity indicators
                    # Look 30 chars before the match for "qt", "quart", "liter", "gal", etc.
                    context_before = sentence_lower[max(0, m.start() - 30):m.start()]
                    capacity_indicators = ['qt', 'quart', 'liter', 'litre', 'gal', 'gallon', 'ml', 'cc']
                    if any(cap_ind in context_before for cap_ind in capacity_indicators) or is_capacity_conversion_engine_match(sentence_lower, m) or is_parenthesized_capacity_conversion(sentence_lower, m):
                        # This is likely a capacity value, not an engine - skip it
                        continue
                    
                    # Look for variant (turbo, v8, etc)
                    variant = ""
                    window_start = max(0, m.start() - 50)
                    window_end = min(len(sentence_lower), m.end() + 100)
                    window = sentence_lower[window_start:window_end]
                    
                    if "turbo" in window and "turbo" not in eng_str.lower():
                        variant = " Turbo"
                    elif "supercharged" in window and "super" not in eng_str.lower():
                        variant = " Supercharged"
                    elif "v6" in window and "v6" not in eng_str.lower():
                        variant = " V6"
                    elif "v8" in window and "v8" not in eng_str.lower():
                        variant = " V8"
                    elif "i4" in window and "i4" not in eng_str.lower():
                        variant = " I4"
                    elif "i6" in window and "i6" not in eng_str.lower():
                        variant = " I6"
                    
                    full_eng = eng_str + variant
                    if full_eng not in engines:
                        engines.append(full_eng)
            except (ValueError, TypeError):
                continue
    
    # No static code-to-displacement fallback. Keep only engines found in document text.
    return list(set(engines))


def extract_engines_from_spec_table(text):
    """
    DYNAMIC TABLE PARSER: Extract engines from specification tables.
    
    Detects engine specification tables (with headers like "Engine", "VIN Code", "Transaxle", etc.)
    and extracts engines directly from table rows with high accuracy.
    
    This is MORE ACCURATE than general sentence parsing for structured data.
    
    Strategy:
    1. Detect specification table headers (look for "Engine" as a column header)
    2. Extract rows following the header
    3. For each row, extract engine specs from the FIRST column (Engine column)
    4. Prioritize exact matches from tables over noisy general extraction
    5. Return engines found in structured table format
    
    Args:
        text: Full PDF text
    
    Returns:
        List of engine strings extracted from tables (e.g., ["3.4L V6", "5.0L V8"])
    """
    table_engines = []
    lines = text.split('\n')
    
    # PASS 1: Find specification table headers
    # Look for lines that contain "Engine" as a column header (often with "VIN", "Transaxle", "Spark")
    # RELAXED MATCHING: Engine might appear without all other keywords on same line
    spec_header_indices = []
    for idx, line in enumerate(lines):
        line_lower = line.lower().strip()
        
        if is_capacity_or_fluid_row(line_lower):
            continue

        # Check if this line is an engine specification/table header.
        # Keep this strict so "Engine Oil with Filter" is not treated as a spec table.
        has_engine_spec_title = "engine" in line_lower and ("spec" in line_lower or "data" in line_lower)
        has_engine_table_header = "engine" in line_lower and any(
            kw in line_lower for kw in ['vin', 'transaxle', 'spark', 'code', 'gap']
        )
        
        # Check for VIN, Transaxle, Spark as indicators of a spec table
        has_table_keywords = any(kw in line_lower for kw in ['vin', 'transaxle', 'spark', 'code', 'gap', 'cubic inches', 'firing order', 'compression ratio'])
        
        if has_engine_spec_title or has_engine_table_header or (has_table_keywords and len(spec_header_indices) < 10):
            spec_header_indices.append(idx)
    
    # PASS 2: For each table header found, extract engines from following rows
    for header_idx in spec_header_indices:
        # Check next 20 lines for engine data (table rows)
        for row_idx in range(header_idx + 1, min(header_idx + 20, len(lines))):
            row = lines[row_idx].strip()
            
            if not row or len(row) < 3:
                continue
            
            # Skip if row looks like another header or separator
            if any(sep in row for sep in ['---', '===', '***']):
                continue
            
            row_lower = row.lower()

            if is_capacity_or_fluid_row(row_lower):
                continue
            
            # CRITICAL: Must contain engine signals AND avoid false positives
            has_multi_engine_header = row_lower.startswith("engine ") and len(list(re.finditer(ENGINE_PATTERN, row_lower))) >= 2
            if not has_engine_signal(row_lower) and not has_multi_engine_header:
                continue
            
            # Skip rows that are clearly about other specs (fuel tank, cargo, towing)
            if any(kw in row_lower for kw in ['cargo', 'towing', 'payload', 'fuel tank', 'weight', 'gvwr']):
                continue
            
            # Extract engines from this row using STRICT ENGINE_PATTERN
            # This captures engines in table rows accurately
            engines_in_row = []
            row_engine_matches = list(re.finditer(ENGINE_PATTERN, row_lower))
            for match_idx, eng_match in enumerate(row_engine_matches):
                try:
                    # CRITICAL: Skip if this match is preceded by capacity indicators
                    context_before = row_lower[max(0, eng_match.start() - 30):eng_match.start()]
                    capacity_indicators = ['qt', 'quart', 'liter', 'litre', 'gal', 'gallon', 'ml', 'cc']
                    if any(cap_ind in context_before for cap_ind in capacity_indicators) or is_capacity_conversion_engine_match(row_lower, eng_match) or is_parenthesized_capacity_conversion(row_lower, eng_match):
                        # This is a capacity value, not an engine - skip it
                        continue
                    
                    eng_size_num = float(eng_match.group(1))
                    if 0.8 <= eng_size_num <= 8.0:
                        eng_str = f"{eng_size_num:.1f}L"
                        
                        # Look for engine variant in this engine's cell/segment, not the next engine column.
                        next_engine_start = row_engine_matches[match_idx + 1].start() if match_idx + 1 < len(row_engine_matches) else len(row_lower)
                        row_window = row_lower[max(0, eng_match.start() - 15):next_engine_start]
                        variant = ""
                        if "v8" in row_window and "v8" not in eng_str.lower():
                            variant = " V8"
                        elif "v6" in row_window and "v6" not in eng_str.lower():
                            variant = " V6"
                        elif "v5" in row_window and "v5" not in eng_str.lower():
                            variant = " V5"
                        elif "i4" in row_window and "i4" not in eng_str.lower():
                            variant = " I4"
                        elif "i6" in row_window and "i6" not in eng_str.lower():
                            variant = " I6"
                        elif "turbo" in row_window and "turbo" not in eng_str.lower():
                            variant = " Turbo"
                        elif "supercharged" in row_window and "super" not in eng_str.lower():
                            variant = " Supercharged"
                        
                        full_eng = eng_str + variant
                        if full_eng not in engines_in_row:
                            engines_in_row.append(full_eng)
                except (ValueError, TypeError):
                    continue
            
            # Add found engines (preferring table-extracted engines)
            table_engines.extend(engines_in_row)
    
    return list(set(table_engines))  # Remove duplicates


def filter_engine_outliers(engines):
    """
    Aggressively remove unlikely engine sizes that are false positives.
    Handles engine variants like "1.4L Turbo", "5.3L V8", etc.
    
    Strategy:
    1. Most vehicles have 1-3 engine options, rarely 4
    2. Remove engines < 1.3L (too small for normal vehicles)
    3. Remove engines > 6.4L (too large for consumer vehicles)
    4. If ANY large gap (>1.2L) exists between engines, remove the smaller isolated group
    5. Maximum 3 engines for most vehicles (compact cars: 2, trucks: 2-3, luxury: 3)
    
    Args:
        engines: List of engine size strings like ["1.4L", "1.4L Turbo", "2.5L", "5.4L V8"]
    
    Returns:
        Filtered list of realistic engines
    """
    if not engines:
        return engines
    
    # If only 1-2 engines, likely legit
    if len(engines) <= 2:
        return engines
    
    # Convert to floats for analysis, extracting numeric part from variants
    engine_nums = []
    for eng in engines:
        try:
            # Extract just the numeric part (e.g., "1.4" from "1.4L Turbo")
            match = re.search(r'(\d+\.?\d*)', eng)
            if match:
                disp = float(match.group(1))
                # More aggressive filtering: only 1.3L-6.4L range
                if 1.3 <= disp <= 6.4:
                    engine_nums.append((disp, eng))
        except (ValueError, AttributeError):
            continue
    
    if not engine_nums:
        return engines
    
    # Sort by displacement
    engine_nums.sort()
    
    # AGGRESSIVE: If we have 4+ engines, definitely has noise
    # Find gaps and identify main clusters
    if len(engine_nums) >= 4:
        gaps = []
        for i in range(len(engine_nums) - 1):
            gap = engine_nums[i + 1][0] - engine_nums[i][0]
            gaps.append((gap, i))
        
        if gaps:
            # Find ANY significant gap (>1.2L) 
            for gap_size, gap_idx in sorted(gaps, reverse=True):
                if gap_size > 1.2:
                    # Split at this gap
                    lower_group = engine_nums[:gap_idx + 1]
                    upper_group = engine_nums[gap_idx + 1:]
                    
                    # Keep the larger group or the more plausible one
                    # For 6.0L+ engines in small car, remove them
                    lower_has_large = any(d > 5.5 for d, _ in lower_group)
                    upper_has_large = any(d > 5.5 for d, _ in upper_group)
                    lower_has_small = any(d < 1.8 for d, _ in lower_group)
                    upper_has_small = any(d < 1.8 for d, _ in upper_group)
                    
                    # If one group is all small (compact) and other is all large (truck), pick one
                    if lower_has_small and not upper_has_small and upper_has_large:
                        # Lower group is compact engines, upper is large
                        engine_nums = lower_group
                        break
                    elif upper_has_small and not lower_has_small and lower_has_large:
                        # Upper group is compact, lower is large
                        engine_nums = upper_group
                        break
                    elif len(lower_group) >= len(upper_group):
                        engine_nums = lower_group
                        break
                    else:
                        engine_nums = upper_group
                        break
    
    # Maximum 3 engines for final result, but keep unique displacements (not just first 3)
    # Group by base displacement, keeping only one variant per displacement
    displacement_map = {}
    for disp, eng_str in engine_nums:
        base_disp = round(disp, 1)  # Group 5.3L and 5.3L V8 together
        
        if base_disp not in displacement_map:
            displacement_map[base_disp] = eng_str
        else:
            # Prefer variants with engine type (V6, V8, etc) over plain displacement
            # Also prefer V8/I4 over V6
            current = displacement_map[base_disp]
            current_has_type = any(t in current for t in ['V6', 'V8', 'I4', 'I3', 'I5', 'I6', 'I8', 'Turbo', 'Supercharged'])
            new_has_type = any(t in eng_str for t in ['V6', 'V8', 'I4', 'I3', 'I5', 'I6', 'I8', 'Turbo', 'Supercharged'])
            
            if new_has_type and not current_has_type:
                # New has type, current doesn't - prefer new
                displacement_map[base_disp] = eng_str
            elif new_has_type and current_has_type:
                # Both have types - prefer V8/I4 over V6
                if ('V8' in eng_str or 'I4' in eng_str) and not ('V8' in current or 'I4' in current):
                    displacement_map[base_disp] = eng_str
    
    # Keep up to 6 unique displacements (increased from 3 to handle trucks with more options)
    result = list(displacement_map.values())[:6]
    
    # Post-processing: Remove 3.8L if it appears with larger legitimate engines (likely fuel context false positive)
    if result and any(eng for eng in result if '3.8' in str(eng)):
        has_larger_v8 = any(eng for eng in result if any(v8_size in str(eng) for v8_size in ['5.4', '6.0', '6.2', '7.0', '5.0', '5.5']))
        if has_larger_v8:
            # Remove 3.8L as it's likely a false positive
            result = [eng for eng in result if '3.8' not in str(eng)]
    
    # If filtering removed everything, return original (fallback)
    return result if result else engines


def consolidate_engine_variants(engines):
    """
    Remove base engine sizes when variant versions exist.
    Example: ["5.3L", "5.3L V8", "6.0L", "6.0L V8"] -> ["5.3L V8", "6.0L V8"]
    
    Strategy:
    1. Group engines by base displacement (e.g., "5.3L")
    2. If a group has both base and variant (e.g., "5.3L" and "5.3L V8")
    3. Keep only the variant versions
    
    Args:
        engines: List of engine strings like ["5.3L", "5.3L V8", "6.0L"]
    
    Returns:
        Consolidated list with only variants when variants exist
    """
    if not engines:
        return engines
    
    # Group engines by base displacement
    groups = {}
    for eng in engines:
        # Extract base displacement (e.g., "5.3L" from "5.3L V8")
        match = re.search(r'(\d+\.?\d*L)', eng)
        if match:
            base = match.group(1)
            if base not in groups:
                groups[base] = []
            groups[base].append(eng)
    
    # For each group, keep only variants if both base and variant exist
    result = []
    for base, eng_list in groups.items():
        if len(eng_list) > 1:
            # Multiple engines with same base - keep variants only
            variants = [e for e in eng_list if e != base]
            if variants:
                result.extend(variants)
            else:
                result.extend(eng_list)
        else:
            # Only one engine with this base
            result.extend(eng_list)
    
    return result


def has_engine_context(text, match_pos, match_text, context_window=150):
    """
    Validate engine type match appears in valid engine context.
    Prevents false positives from page numbers, codes, etc.
    
    Step-by-step process:
    1. Define engine-related keywords to check for
    2. Extract text window around match position
    3. For F8 (rare engine), require strict validation with "engine"/"type"/"spec"
    4. For other types, check for any engine keyword in context
    5. Return True if valid context found, False otherwise
    
    Args:
        text: Full PDF text (lowercase)
        match_pos: Character position of match start
        match_text: The matched text string
        context_window: Number of chars to check before/after match
    
    Returns:
        Boolean indicating if match has valid context
    """
    engine_keywords = [
        "engine", "displacement", "cc", "cylinder", "oil", "turbo",
        "configuration", "specs", "specifications", "performance",
        "l ", " l", "liter", "litre", "capacity", "horsepower", "hp"
    ]
    
    start = max(0, match_pos - context_window)
    end = min(len(text), match_pos + len(match_text) + context_window)
    context = text[start:end].lower()
    
    if match_text.lower() == "f8":
        close_context = text[max(0, match_pos - 50):min(len(text), match_pos + len(match_text) + 50)].lower()
        strict_keywords = ["engine", "type", "spec", "f8"]
        requires_engine = any(kw in close_context for kw in ["engine", "engine type", "engine configuration"])
        if not requires_engine:
            return False
    
    has_context = any(keyword in context for keyword in engine_keywords)
    return has_context


def extract_engine_types(text, all_engines=None):
    """
    Extract engine types from actual engine descriptions.
    DYNAMIC FALLBACK: If no explicit types found, infer from ACTUAL ENGINE displacements (not all PDF mentions).
    
    Step-by-step process:
    1. Extract explicit types from engine names (e.g., "1.4L Turbo" → "TURBO")
    2. If engine_types is empty, use DYNAMIC fallback based on displacement of EXTRACTED engines only
    3. Fallback rules are universal (work for ANY make/model/year):
       - Uses only displacements from actual engines found, not references in PDF text
       - <1.6L → I3, I4
       - 1.6-2.5L → I4
       - 2.5-3.0L → I4
       - 3.0-3.5L → V6
       - 3.5-5.0L → V6, V8
       - 5.0L+ → V8
    4. Validate PASS 1 results against actual engine displacements (remove false positives)
    5. Normalize and deduplicate
    
    Args:
        text: Full PDF text
        all_engines: Optional list of extracted engine names (e.g., ['1.4L', '1.8L', '2.4L Turbo'])
    
    Returns:
        List of unique engine types found (from explicit specs or dynamic inference)
    """
    engine_types_found = set()
    text_lower = text.lower()
    lines = text_lower.split('\n')
    
    # PASS 1: Extract explicit engine types from engine descriptions
    for line in lines:
        # Find all engine size matches in this line
        engine_matches = list(re.finditer(ENGINE_PATTERN, line))
        if not engine_matches:
            continue
        
        # For each engine size found, extract descriptor text around it
        for engine_match in engine_matches:
            engine_start = engine_match.start()
            engine_end = engine_match.end()
            engine_size_text = engine_match.group()
            
            # Extract nearby text (up to next comma, semicolon, or 40 chars)
            end_search = min(len(line), engine_end + 50)
            trailing_text = line[engine_end:end_search]
            
            # Extract up to next delimiter
            delimiter_pos = len(trailing_text)
            for delim in [',', ';', '(', ')']:
                pos = trailing_text.find(delim)
                if pos != -1 and pos < delimiter_pos:
                    delimiter_pos = pos
            
            descriptor_text = (engine_size_text + trailing_text[:delimiter_pos]).lower()
            
            # Extract explicit engine types from descriptor
            mechanical_pattern = r"\b(v3|v4|v5|v6|v8|v10|v12|v16|v20|v24|i3|i4|i5|i6|i8|h4|w8|w12|w16|f8|flat[-]?4|flat[-]?6|flat[-]?12|boxer|boxe|rotary|wankel|turbo|dual[-]?turbo|quad[-]?turbo|sequential[-]?turbo|supercharged|twin[-]?supercharged|twin[-]?turbo|twin[-]?scroll)\b"
            
            for match in re.finditer(mechanical_pattern, descriptor_text):
                match_text = match.group().strip().replace(" ", "").upper()
                engine_types_found.add(match_text)
            
            # Inline/cylinder patterns
            inline_pattern = r"\binline\s*[-]?\s*([3-8])\b"
            for match in re.finditer(inline_pattern, descriptor_text):
                digit = match.group(1)
                engine_types_found.add(f"I{digit}")
            
            cylinder_pattern = r"\b([3-8])\s*[-]?\s*cylinder\b"
            for match in re.finditer(cylinder_pattern, descriptor_text):
                digit = match.group(1)
                engine_types_found.add(f"I{digit}")
    
    # PASS 2: Dynamic fallback + validation - refine inference using ACTUAL ENGINE displacements
    # Extract displacements ONLY from the actual engines found (not all PDF mentions)
    # This prevents false positives from comparison/reference engine sizes
    engine_displacements = []
    if all_engines:
        for eng in all_engines:
            match = re.search(r'(\d+\.?\d*)', eng)
            if match:
                try:
                    displacement = float(match.group(1))
                    engine_displacements.append(displacement)
                except (ValueError, AttributeError):
                    pass
    
    if engine_displacements:
        min_disp = min(engine_displacements)
        max_disp = max(engine_displacements)
        
        # If PASS 1 found nothing, populate from displacement
        if not engine_types_found:
            if max_disp < 1.3:
                engine_types_found.add("I3")
            elif max_disp < 2.5:
                engine_types_found.add("I4")
            elif max_disp < 3.0:
                engine_types_found.add("I4")
            elif max_disp < 3.5:
                engine_types_found.add("V6")
            elif max_disp < 5.0:
                engine_types_found.add("V6")
                engine_types_found.add("V8")
            else:
                engine_types_found.add("V8")
        else:
            # PASS 1 found types - validate against actual engine displacements
            # If found V8/V10/V12 but displacement < 3.0L, likely false positive
            invalid_types = set()
            for et in engine_types_found:
                if et in ("V8", "V10", "V12", "V16") and max_disp < 3.0:
                    invalid_types.add(et)  # Too large for small displacement
                elif et in ("I3", "I4", "I5", "I6") and min_disp > 4.0:
                    invalid_types.add(et)  # Too small for large displacement
            
            # Remove invalid types and replace with displacement-based inference
            if invalid_types:
                engine_types_found -= invalid_types
                
                # If we removed everything, fill in from displacement
                if not engine_types_found:
                    if max_disp < 1.3:
                        engine_types_found.add("I3")
                    elif max_disp < 2.5:
                        engine_types_found.add("I4")
                    elif max_disp < 3.0:
                        engine_types_found.add("I4")
                    elif max_disp < 3.5:
                        engine_types_found.add("V6")
                    elif max_disp < 5.0:
                        engine_types_found.add("V6")
                        engine_types_found.add("V8")
                    else:
                        engine_types_found.add("V8")
    
    # Normalize aliases
    normalized_types = set()
    for engine_type in engine_types_found:
        if engine_type in ("BOXE", "BOXER"):
            normalized_types.add("BOXER")
        elif engine_type.startswith("FLAT"):
            normalized_types.add("FLAT-6")
        elif engine_type in ("TWINTTURBO", "TWIN-TURBO"):
            normalized_types.add("TWIN-TURBO")
        elif engine_type == "BITURBO":
            normalized_types.add("BI-TURBO")
        else:
            normalized_types.add(engine_type)
    
    return sorted(list(normalized_types))  # Return sorted list for consistency


def engine_matches_capacity(engine, capacity):
    """
    Validate that engine size matches expected oil capacity.
    Typical small engines use 3-4 qts, medium 4-5.5, large 5+ qts.
    
    Handles engine variants like "1.4L Turbo" or "5.3L V8".
    
    Args:
        engine: Engine size string (e.g., "1.4L" or "1.4L Turbo")
        capacity: Capacity dict with "quarts" key
    
    Returns:
        Boolean indicating if engine/capacity pair makes sense
    """
    if not capacity:
        return True
    
    quarts = capacity.get("quarts", 0)
    
    # Extract numeric part from engine (e.g., "1.4L Turbo" → "1.4")
    match = re.search(r'(\d+\.?\d*)', engine)
    if not match:
        return True
    
    try:
        size = float(match.group(1))
    except (ValueError, AttributeError):
        return True
    
    if size <= 1.6 and 3 <= quarts <= 4:
        return True
    if 1.7 <= size <= 2.5 and 4 <= quarts <= 5.5:
        return True
    if size >= 2.6 and quarts >= 5:
        return True
    
    return False


def extract_temperature(sentence):
    """
    Extract temperature values and classify conditions DYNAMICALLY.
    Converts Celsius to Fahrenheit and intelligently classifies based on temp ranges.
    
    SMART LOGIC:
    - If BOTH cold (<40°F) AND hot (>85°F) temps found → "all temperatures" only (don't list individual temps)
    - If ONLY cold temps (≤40°F) → "cold weather" + range or specific temps
    - If ONLY hot temps (≥85°F) → "hot weather" + range or specific temps
    - If moderate range (40-85°F) → specific temps or range
    
    Args:
        sentence: Text to search for temperature data
    
    Returns:
        Set of temperature information strings (dynamically determined)
    """
    s = sentence.lower()
    temps = re.findall(TEMP_PATTERN, s)
    values = []
    
    for value, unit in temps:
        value = int(value)
        if unit == "c":
            value = round((value * 9 / 5) + 32)
        # FILTER: Only keep realistic automotive temperatures (-60°F to 150°F)
        if -60 <= value <= 150:
            values.append(value)
    
    result = set()
    
    if not values:
        return {"all temperatures"}
    
    min_temp, max_temp = min(values), max(values)
    
    # SMART CLASSIFICATION: Determine the condition type
    has_cold = any(t <= 40 for t in values)    # Cold: ≤40°F
    has_hot = any(t >= 85 for t in values)     # Hot: ≥85°F
    
    if has_cold and has_hot:
        # BOTH cold and hot temps → covers all temperatures
        # DO NOT list individual temps - just say "all temperatures"
        result.add("all temperatures")
        # Only include range if it's meaningful (not the full -60 to 150)
        if not (min_temp <= -50 and max_temp >= 140):  # Not the extreme full range
            result.add(f"range: {min_temp}F to {max_temp}F")
    elif has_cold:
        # ONLY cold temps → focus on cold weather with range
        result.add("cold weather")
        if len(values) > 1:
            result.add(f"range: {min_temp}F to {max_temp}F")
        else:
            result.add(f"{values[0]}F")
    elif has_hot:
        # ONLY hot temps → focus on hot weather with range
        result.add("hot weather")
        if len(values) > 1:
            result.add(f"range: {min_temp}F to {max_temp}F")
        else:
            result.add(f"{values[0]}F")
    else:
        # Moderate range temps (40-85°F) - return specific temps if few, else range
        if len(values) <= 3:
            for temp in values:
                result.add(f"{temp}F")
        else:
            result.add(f"range: {min_temp}F to {max_temp}F")
    
    return result


def normalize_capacity_value(quarts_value):
    """
    Intelligently normalize capacity values - round near-whole numbers.
    E.g., 6.98 → 7.0, 4.02 → 4.0, but 6.34 → 6.34
    
    Strategy:
    1. If value is within 0.05 of a whole number, round it
    2. Otherwise, keep original precision
    3. Helps catch values like 6.98 (clearly meant to be 7.0)
    
    Args:
        quarts_value: Numeric capacity in quarts
    
    Returns:
        Normalized quarts value
    """
    if not isinstance(quarts_value, (int, float)):
        return quarts_value
    
    rounded = round(quarts_value)
    if abs(quarts_value - rounded) < 0.05:
        return float(rounded)
    return round(quarts_value, 2)


def find_correct_engine_oil_capacities(doc):
    """
    Search PDF for correct engine oil capacities by looking forlines with "Engine Oil" keyword.
    
    Strategy:
    1. Find lines that explicitly contain "Engine Oil" or "engine oil with filter"
    2. Extract engines and capacities from these SPECIFIC lines
    3. Validate against document-derived plausible capacities
    4. Return dict of engine -> correct capacity
    
    Args:
        doc: PyMuPDF document object
    
    Returns:
        Dict mapping engine sizes to correct capacities
    """
    correct_capacities = {}
    
    for page_num, page in enumerate(doc):
        text = page.get_text()
        text_lower = text.lower()
        
        # Only search specification pages
        if not ("capacity" in text_lower or "specification" in text_lower):
            continue
        
        lines = text.split('\n')
        
        for line_idx, line in enumerate(lines):
            line_lower = line.lower()
            
            # STRICT: Line MUST contain "engine oil" to be considered
            if "engine oil" not in line_lower:
                continue
            
            # Skip if it mentions other fluids
            if "coolant" in line_lower or "transmission" in line_lower or "cooling" in line_lower:
                continue
            
            # Find engines and capacities in this line
            engines = [(float(m.group(1)), m.start()) for m in re.finditer(ENGINE_PATTERN, line_lower)]
            capacities = [(m.group(1), m.group(2), m.start()) for m in re.finditer(CAPACITY_PATTERN, line_lower, re.IGNORECASE)]
            
            if engines and capacities:
                # Engines and capacities both in same line
                for eng_val, eng_pos in engines:
                    eng_size = f"{eng_val:.1f}L"
                    
                    # Find nearest capacity after engine
                    for cap_str, cap_unit, cap_pos in capacities:
                        if cap_pos > eng_pos:
                            try:
                                q, l = to_quarts_liters(cap_str, cap_unit)
                                if 1.0 <= q <= 12.0:
                                    correct_capacities[eng_size] = {
                                        "quarts": round(q, 2),
                                        "liters": round(l, 1)
                                    }
                                    break
                            except (ValueError, TypeError):
                                continue
            else:
                # Maybe separate formatting: engine on current line, capacity on next
                if engines and line_idx + 1 < len(lines):
                    next_line = lines[line_idx + 1]
                    next_lower = next_line.lower()
                    next_capacities = [(m.group(1), m.group(2), m.start()) for m in re.finditer(CAPACITY_PATTERN, next_lower, re.IGNORECASE)]
                    
                    # Only use next line capacities if they're not on an "engine oil" line (avoid duplicates)
                    if next_capacities and "engine oil" not in next_lower:
                        for eng_val, _ in engines:
                            eng_size = f"{eng_val:.1f}L"
                            
                            for cap_str, cap_unit, _ in next_capacities:
                                try:
                                    q, l = to_quarts_liters(cap_str, cap_unit)
                                    if 1.0 <= q <= 12.0:
                                        correct_capacities[eng_size] = {
                                            "quarts": round(q, 2),
                                            "liters": round(l, 1)
                                        }
                                        break
                                except (ValueError, TypeError):
                                    continue
    
    return correct_capacities


def apply_correct_capacities(multi_engine_data, doc):
    """
    DYNAMIC: Validate engine oil capacities using PDF data only.
    No hardcoded ranges - uses actual extracted capacities to validate.
    
    Strategy:
    1. For each engine, extract ALL capacities mentioned in PDF for that engine
    2. Use the most common or most reasonable capacity
    3. If significantly different from current, update it
    4. No hardcoded validation - only PDF data matters
    
    Args:
        multi_engine_data: Extracted engine data to validate/correct
        doc: PyMuPDF document object
    
    Returns:
        Corrected multi_engine_data
    """
    if not multi_engine_data:
        return multi_engine_data
    
    # For each engine, dynamically extract all capacities from PDF
    for eng_str in list(multi_engine_data.keys()):
        cap_info = multi_engine_data[eng_str].get("oil_capacity", {})
        with_filter = cap_info.get("with_filter")
        
        if not with_filter or "quarts" not in with_filter:
            continue
        
        current_q = with_filter["quarts"]
        
        # DYNAMIC: Extract all capacities from PDF for this engine
        all_capacities = extract_all_capacities_for_engine(doc, eng_str)
        
        if not all_capacities:
            # No additional capacities found, keep current
            continue
        
        # Use most common capacity, or if tie, use median
        from collections import Counter
        capacity_counts = Counter([round(c, 1) for c in all_capacities])
        most_common_cap = capacity_counts.most_common(1)[0][0]
        
        # If current capacity is significantly different, update it
        # (using 0.5 quart tolerance for rounding differences)
        if abs(current_q - most_common_cap) > 0.5:
            # Convert back to liters for the update
            new_liters = round(most_common_cap * 0.946, 1)
            multi_engine_data[eng_str]["oil_capacity"]["with_filter"] = {
                "quarts": round(most_common_cap, 2),
                "liters": new_liters
            }
            print(f"      CORRECTED: {eng_str} capacity {current_q}qt → {round(most_common_cap, 2)}qt")
    
    return multi_engine_data


def extract_all_capacities_for_engine(doc, engine_str):
    """
    DYNAMIC: Extract engine OIL capacity for a specific engine from PDF.
    Only looks for explicit "engine oil" capacity mentions - ignores coolant, fuel, transmission.
    
    Strategy:
    1. Search for engine oil capacity sections in PDF
    2. Match by engine size in oil-specific context
    3. Return only engine oil capacities (1.0-12.0 quarts)
    4. Return empty if ambiguous or not found
    
    Args:
        doc: PyMuPDF document object
        engine_str: Engine size like "1.4L", "5.4L V8", etc.
    
    Returns:
        List of OIL capacity values (in quarts) found in PDF for this engine
    """
    capacities = []
    
    # Extract numeric engine size
    match = re.search(r'(\d+\.?\d*)', engine_str)
    if not match:
        return capacities
    
    eng_num = match.group(1)
    
    # Search entire PDF for EXPLICIT engine oil capacity sections
    for page in doc:
        text = page.get_text()
        text_lower = text.lower()
        
        # Look for lines with explicit "oil" indicators and this engine size
        # Patterns that indicate engine oil (not coolant/transmission/fuel):
        # "engine oil", "oil with filter", "oil capacity", "oil drain", "oil change"
        oil_keywords = r"(?:engine\s+oil|oil\s+with\s+filter|oil\s+capacity|oil\s+change|oil\s+drain|oil\s+fill)"
        
        # Find sections that mention both oil keywords AND the engine size
        for match in re.finditer(oil_keywords, text_lower, re.IGNORECASE):
            # Extract window around this oil mention
            start = max(0, match.start() - 50)
            end = min(len(text), match.end() + 200)
            window = text[start:end]
            
            # Check if engine size is in this window
            if eng_num.replace(".", r"\.") in window or eng_num in window.lower():
                # Find all capacities in this window
                for cap_match in re.finditer(CAPACITY_PATTERN, window, re.IGNORECASE):
                    try:
                        q, l = to_quarts_liters(cap_match.group(1), cap_match.group(2))
                        # Only add if it seems reasonable (1-12 quarts for engine oil)
                        if 1.0 <= q <= 12.0:
                            capacities.append(q)
                    except (ValueError, TypeError):
                        continue
    
    return capacities


def detect_all_engines_in_pdf(doc):
    """
    Aggressively extract ALL engine sizes mentioned in PDF.
    Used as fallback when engine-specific extraction fails.
    
    Strategy:
    1. Search entire PDF text for engine patterns (not just capacity pages)
    2. Look for explicit engine mentions in spec tables
    3. Extract with associated engine types (V8, Turbo, etc.)
    4. Return list of all found engines
    
    Args:
        doc: PyMuPDF document object
    
    Returns:
        List of engine size strings found (e.g., ["5.3L", "5.3L V8"])
    """
    all_engines = []
    full_text = ""
    
    # Collect all text from PDF
    for page in doc:
        full_text += page.get_text().lower() + "\n"
    
    # Find all engine mentions with context
    engine_matches = []
    for m in re.finditer(ENGINE_PATTERN, full_text):
        eng_size = f"{float(m.group(1)):.1f}L"
        start = max(0, m.start() - 100)
        end = min(len(full_text), m.end() + 200)
        context = full_text[start:end]
        
        engine_matches.append((eng_size, context, m.start()))
    
    # Extract variants and deduplicate
    seen = set()
    for eng_size, context, pos in engine_matches:
        # Look for engine type in context
        variant = ""
        if re.search(r'\bv8\b', context):
            variant = " V8"
        elif re.search(r'\bv6\b', context):
            variant = " V6"
        elif re.search(r'\bi4\b', context):
            variant = " I4"
        elif re.search(r'\bi6\b', context):
            variant = " I6"
        elif re.search(r'\bturbo\b', context):
            variant = " Turbo"
        elif re.search(r'\bsupercharged\b', context):
            variant = " Supercharged"
        
        full_eng = eng_size + variant
        if full_eng not in seen:
            all_engines.append(full_eng)
            seen.add(full_eng)
    
    return all_engines


def fix_unknown_engine(doc, engine_data):
    """
    Intelligently replace "unknown_engine" with detected actual engines.
    
    Strategy:
    1. Detect all engines mentioned in PDF
    2. Extract engine sizes from capacity table context
    3. Match detected engines to unknown_engine's capacity
    4. Replace unknown_engine with best matched engine
    
    Args:
        doc: PyMuPDF document object
        engine_data: Dict of engine → capacity info
    
    Returns:
        Corrected engine_data dict with unknown_engine replaced if possible
    """
    if "unknown_engine" not in engine_data:
        return engine_data
    
    # Get the unknown capacity for matching
    unknown_cap = engine_data["unknown_engine"]["oil_capacity"]["with_filter"]
    if not unknown_cap:
        return engine_data
    
    # Detect all engines in PDF
    detected_engines = detect_all_engines_in_pdf(doc)
    
    if not detected_engines:
        return engine_data
    
    # Find best matching engine by closest document-observed capacity.
    best_match = None
    best_diff = float("inf")
    cap_q = unknown_cap.get("quarts", 0)
    for eng in detected_engines:
        observed_caps = extract_all_capacities_for_engine(doc, eng)
        if not observed_caps:
            continue
        local_best = min(abs(cap_q - obs) for obs in observed_caps)
        if local_best < best_diff:
            best_diff = local_best
            best_match = eng
    
    # If found a match, replace unknown_engine
    if best_match:
        engine_data[best_match] = engine_data.pop("unknown_engine")
    
    return engine_data


def extract_engine_capacities(doc):
    """
    Extract engine oil capacities ONLY from explicit "Engine Oil" sections.
    
    SIMPLE & STRICT:
    1. Scan for lines containing "Engine Oil with Filter" or "Engine Oil" header
    2. Only extract engines and capacities from the NEXT 10 lines after that header
    3. Match engine-capacity pairs on the same line (table row format)
    4. Ignore all other sections (cooling, transmission, fuel, air conditioning)
    
    Args:
        doc: PyMuPDF document object
    
    Returns:
        Dict mapping engine sizes to capacity objects
    """
    engine_caps = {}
    
    # Collect all PDF text
    full_text = ""
    for page in doc:
        full_text += page.get_text() + "\n"
    
    lines = full_text.split('\n')
    
    # Find all lines with "Engine Oil" keywords and extract ONLY from those sections
    for line_idx, line in enumerate(lines):
        line_lower = line.lower()
        
        # Check if this line is an engine oil header
        is_oil_header = any(kw in line_lower for kw in [
            'engine oil', 'motor oil', 'oil with filter', 'oil capacity', 'engine oil capacity'
        ])
        
        if not is_oil_header:
            continue
        
        # Found an oil header. Now extract from THIS line + next 10 lines
        oil_section_lines = lines[line_idx:min(line_idx + 15, len(lines))]
        
        for section_line_idx, section_line in enumerate(oil_section_lines):
            section_line_lower = section_line.lower()
            
            # Skip empty or very short lines
            if len(section_line.strip()) < 3:
                continue
            
            # CRITICAL: Skip lines mentioning other fluids (cooling, transmission, fuel)
            if any(kw in section_line_lower for kw in ['cooling', 'coolant', 'transmission', 'transaxle', 'differential', 'power steering', 'air conditioning', 'refrigerant', 'fuel tank', 'wheel nut']):
                continue

            capacity_matches = [
                cm for cm in re.finditer(CAPACITY_PATTERN, section_line_lower)
                if is_real_capacity_match(section_line_lower, cm)
            ]
            engine_matches = list(re.finditer(ENGINE_PATTERN, section_line_lower))
            capacity_from_oil_header = False

            if not capacity_matches and not engine_matches:
                if "engine oil" not in section_line_lower and "oil with filter" not in section_line_lower:
                    continue

                for next_line in oil_section_lines[section_line_idx + 1:section_line_idx + 4]:
                    next_line_lower = next_line.lower()
                    if any(kw in next_line_lower for kw in ['cooling', 'coolant', 'transmission', 'transaxle', 'differential', 'power steering', 'air conditioning', 'refrigerant', 'fuel tank', 'wheel nut']):
                        break

                    capacity_matches = [
                        cm for cm in re.finditer(CAPACITY_PATTERN, next_line_lower)
                        if is_real_capacity_match(next_line_lower, cm)
                    ]
                    if capacity_matches:
                        section_line_lower = next_line_lower
                        capacity_from_oil_header = True
                        break

                if not capacity_matches:
                    continue
            
            # CRITICAL FILTER: Skip engine matches that are actually capacity values
            # (i.e., a number followed by qt/quart/liter comes RIGHT BEFORE the engine match)
            capacity_number_pattern = r"\d+\.?\d*\s*(?:qt|quart|qts|gal|gallon|ml|cc)"
            filtered_engine_matches = []
            for em in engine_matches:
                # Check if there's a capacity pattern (number + unit) immediately before this engine match
                context_before = section_line_lower[max(0, em.start() - 20):em.start()]
                if (
                    re.search(r"\d+\.?\d*\s*" + CAPACITY_UNIT_PATTERN + r"\s*\(?\s*$", context_before, re.I)
                    or is_capacity_conversion_engine_match(section_line_lower, em)
                    or is_parenthesized_capacity_conversion(section_line_lower, em)
                ):
                    # This engine match is actually part of a capacity value, skip it
                    continue
                filtered_engine_matches.append(em)

            engine_matches = filtered_engine_matches

            if not engine_matches:
                if not capacity_matches:
                    continue

                if not capacity_from_oil_header and "engine oil" not in section_line_lower and "oil with filter" not in section_line_lower:
                        continue

                selected_cap = None
                for cm in capacity_matches:
                    unit = cm.group(2).lower()
                    if 'qt' in unit or 'quart' in unit:
                        selected_cap = cm
                        break

                if not selected_cap:
                    selected_cap = capacity_matches[0]

                try:
                    q, l = to_quarts_liters(selected_cap.group(1), selected_cap.group(2))
                    if 3.0 <= q <= 12.0 and "unknown_engine" not in engine_caps:
                        engine_caps["unknown_engine"] = {
                            "with_filter": {"quarts": q, "liters": l},
                            "without_filter": None
                        }
                except (ValueError, TypeError, AttributeError):
                    pass

                continue
            
            # Prefer engine token that comes before first capacity token
            first_cap_pos = capacity_matches[0].start() if capacity_matches else len(section_line_lower)
            eng_match = None
            for em in engine_matches:
                if em.start() < first_cap_pos:
                    eng_match = em
                    break
            
            if eng_match is None and engine_matches:
                eng_match = engine_matches[0]
            
            if not eng_match:
                continue
            
            try:
                eng_size = f"{float(eng_match.group(1)):.1f}L"
                
                # Extract the FIRST capacity after the engine (prefer quarts)
                selected_cap = None
                for cm in capacity_matches:
                    if cm.start() >= eng_match.end():
                        # Prefer quart values
                        unit = cm.group(2).lower()
                        if 'qt' in unit or 'quart' in unit:
                            selected_cap = cm
                            break
                
                # If no quart found, take first capacity after engine
                if not selected_cap:
                    for cm in capacity_matches:
                        if cm.start() >= eng_match.end():
                            selected_cap = cm
                            break
                
                if not selected_cap:
                    for next_line in oil_section_lines[section_line_idx + 1:section_line_idx + 4]:
                        next_line_lower = next_line.lower()
                        if any(kw in next_line_lower for kw in ['cooling', 'coolant', 'transmission', 'transaxle', 'differential', 'power steering', 'air conditioning', 'refrigerant', 'fuel tank', 'wheel nut']):
                            break

                        next_capacity_matches = [
                            cm for cm in re.finditer(CAPACITY_PATTERN, next_line_lower)
                            if is_real_capacity_match(next_line_lower, cm)
                        ]
                        if not next_capacity_matches:
                            continue

                        for cm in next_capacity_matches:
                            unit = cm.group(2).lower()
                            if 'qt' in unit or 'quart' in unit:
                                selected_cap = cm
                                break

                        if not selected_cap:
                            selected_cap = next_capacity_matches[0]
                        break

                if not selected_cap:
                    continue
                
                q, l = to_quarts_liters(selected_cap.group(1), selected_cap.group(2))
                
                # Only accept reasonable oil capacity values
                if 3.0 <= q <= 12.0:
                    engine_caps[eng_size] = {
                        "with_filter": {"quarts": q, "liters": l},
                        "without_filter": None
                    }
            except (ValueError, TypeError, AttributeError):
                continue
    
    return engine_caps


def extract_fallback_capacity(doc):
    """
    Extract generic oil capacity as fallback when engine-specific data unavailable.
    Used for documents with general oil info but no engine-specific specs.
    
    Step-by-step process:
    1. Search pages with oil keywords
    2. Priority: Look for "with filter" or "including filter" mentions
    3. Capture capacity value following filter keywords
    4. Fallback: Use first valid capacity found on page
    5. Return with_filter variant
    
    Args:
        doc: PyMuPDF document object
    
    Returns:
        Capacity dict or None if not found
    """
    for page in doc:
        text = page.get_text().lower()
        if not any(kw in text for kw in OIL_PAGE_KEYWORDS):
            continue
        
        wf_m = re.search(
            r"(?:with|including)\s*(?:filter|oil\s*filter)[^0-9]{0,60}" + CAPACITY_PATTERN,
            text
        )
        if wf_m:
            window = text[wf_m.start():min(len(text), wf_m.end() + 80)]
            capacity_matches = [
                m for m in re.finditer(CAPACITY_PATTERN, window)
                if is_real_capacity_match(window, m)
            ]
            selected = None
            for m in capacity_matches:
                unit = m.group(2).lower()
                if 'qt' in unit or 'quart' in unit:
                    selected = m
                    break
            if not selected:
                selected = wf_m

            q, l = to_quarts_liters(selected.group(1), selected.group(2))
            if 3.0 <= q <= 9.0:
                return {"with_filter": {"quarts": q, "liters": l}, "without_filter": None}
        
        for m in re.finditer(CAPACITY_PATTERN, text):
            if not is_real_capacity_match(text, m):
                continue
            q, l = to_quarts_liters(m.group(1), m.group(2))
            if 3.0 <= q <= 9.0:
                return {"with_filter": {"quarts": q, "liters": l}, "without_filter": None}
    
    return None

def extract_oils(text):
    """
    Extract oil types with recommendation strength and temperature ranges.
    Filters out "Do not use" oils and prioritizes "best viscosity grade" recommendations.
    
    Returns:
        Tuple of (oil_scores dict, oil_temps dict, engine_oil_map_inline dict)
        where engine_oil_map_inline maps engine sizes to lists of oils
    """
    oil_scores = {}
    oil_temps = {}
    engine_oil_map_inline = {}  # Map of engine -> [oils] (list of oils for each engine)
    do_not_use_oils = set()
    
    # PASS 0: Identify "Do not use" oils - these should be filtered out
    # Be precise: only extract oils mentioned as explicitly unsuitable, not secondary recommendations
    do_not_use_pattern = r"do\s+not\s+use|should\s+not\s+be\s+used"
    for match in re.finditer(do_not_use_pattern, text.lower()):
        start_pos = match.start()
        
        # Find the sentence boundary (Period + uppercase or newline + uppercase or "SAE")
        end_pos = start_pos + 500
        for i in range(match.end(), min(len(text), match.end() + 500)):
            if (text[i] == '.' or text[i] == '\n') and i + 1 < len(text):
                j = i + 1
                while j < len(text) and text[j] in ' \n\t':
                    j += 1
                # Stop at sentence boundary
                if j < len(text) and (text[j].isupper() or text[j:j+3].upper() == 'SAE'):
                    end_pos = i
                    break
        
        window = text[start_pos:end_pos]
        
        # Skip if this is clearly a "may be used" alternative context
        if "may be used" in window.lower() or "can be used" in window.lower() or "alternative" in window.lower():
            continue
            
        for oil_match in re.finditer(OIL_PATTERN, window.lower(), re.I):
            if oil_match:
                base, grade = oil_match.groups()[-2:]
                oil = normalize_oil(f"{base}W-{grade}")
                do_not_use_oils.add(oil)
    
    # PASS 1: Extract inline oils from engine spec lines
    # DYNAMIC: Extract actual temperature context around the oil mention, don't hardcode
    lines = text.split('\n')
    for line in lines:
        line_lower = line.lower()
        if re.search(ENGINE_PATTERN, line_lower) and '(sae' in line_lower:
            eng_match = re.search(ENGINE_PATTERN, line_lower)
            if eng_match:
                eng_size = f"{float(eng_match.group(1)):.1f}L"
                oil_match = re.search(r'\(sae\s+' + OIL_PATTERN, line_lower, re.I)
                if oil_match:
                    base, grade = oil_match.groups()[-2:]
                    oil = normalize_oil(f"{base}W-{grade}")
                    if oil not in do_not_use_oils:
                        if oil not in oil_scores:
                            oil_scores[oil] = 0
                            oil_temps[oil] = set()
                        oil_scores[oil] += 4
                        if eng_size not in engine_oil_map_inline:
                            engine_oil_map_inline[eng_size] = []
                        if oil not in engine_oil_map_inline[eng_size]:
                            engine_oil_map_inline[eng_size].append(oil)
                        # Use extracted temperatures first with a generic fallback when missing.
                        temps = extract_temperature(line)
                        final_temps = get_temperature_with_fallback(temps, oil)
                        oil_temps[oil].update(final_temps)
    
    # PASS 2: Find "best viscosity grade" statements - these get VERY HIGH priority
    best_oil_pattern = r'sae\s+' + OIL_PATTERN + r'(?:\s+is\s+(?:the\s+)?best|is\s+(?:the\s+)?best\s+viscosity)'
    for match in re.finditer(best_oil_pattern, text.lower(), re.I):
        if match:
            base, grade = match.groups()[-2:]
            oil = normalize_oil(f"{base}W-{grade}")
            if oil not in do_not_use_oils:
                if oil not in oil_scores:
                    oil_scores[oil] = 0
                    oil_temps[oil] = set()
                oil_scores[oil] += 10
                
                # Check for engine-specific recommendations
                # Extract only the FIRST sentence/statement from the oil match
                # This prevents capturing multiple oil recommendations when they're close together
                sentence_start = match.start()
                sentence_end = min(len(text), match.end() + 300)
                
                # Find the sentence boundary (. or \n followed by an uppercase letter or "SAE")
                for i in range(match.end(), min(len(text), match.end() + 300)):
                    if (text[i] == '.' or text[i] in '\n') and i + 1 < len(text):
                        # Check if next non-space char is uppercase or "SAE"
                        j = i + 1
                        while j < len(text) and text[j] in ' \n\t':
                            j += 1
                        if j < len(text) and (text[j].isupper() or text[j:j+3].upper() == 'SAE'):
                            sentence_end = i
                            break
                
                context = text[sentence_start:sentence_end].lower()
                
                # Extract engines from this statement ONLY
                all_engines_in_context = []
                for eng_m in re.finditer(ENGINE_PATTERN, context):
                    eng_size = f"{float(eng_m.group(1)):.1f}L"
                    all_engines_in_context.append(eng_size)
                    if eng_size not in engine_oil_map_inline:
                        engine_oil_map_inline[eng_size] = []
                    if oil not in engine_oil_map_inline[eng_size]:
                        engine_oil_map_inline[eng_size].append(oil)
                
                if not oil_temps[oil]:
                    oil_temps[oil].add("all temperatures")

    # PASS 2.25: Capture generic preferred/recommended engine-oil statements.
    # This is intentionally dynamic and phrase-based rather than vehicle-specific.
    preferred_patterns = [
        r'(?:an?\s+oil\s+with\s+a\s+viscosity\s+of\s+)' + OIL_PATTERN + r'[^.]{0,180}?\b(?:is\s+preferred|preferred|recommended)\b',
        r'(?:recommended\s+engine\s+oil[^.]{0,160}?)(?:sae\s+)?' + OIL_PATTERN,
        r'(?:engine\s+oil[^.]{0,120}?)(?:sae\s+)?' + OIL_PATTERN + r'[^.]{0,120}?\b(?:preferred|recommended|viscosity)\b',
    ]

    for pattern in preferred_patterns:
        for match in re.finditer(pattern, text.lower(), re.I):
            base, grade = match.groups()[-2:]
            oil = normalize_oil(f"{base}W-{grade}")
            if oil in do_not_use_oils:
                continue

            context_start = max(0, match.start() - 80)
            context_end = min(len(text), match.end() + 120)
            context = text[context_start:context_end]

            if has_non_engine_oil_context(context) or not has_engine_oil_context(context):
                continue

            if oil not in oil_scores:
                oil_scores[oil] = 0
                oil_temps[oil] = set()

            oil_scores[oil] += 6
            temps = extract_temperature(context)
            final_temps = get_temperature_with_fallback(temps, oil)
            oil_temps[oil].update(final_temps)
    
    # PASS 2.5: Extract alternative/conditional oils (e.g., "0W-30 may be used in extreme cold")
    # These are typically mentioned in "Cold Temperature Operation" or similar sections
    # Use a simplified pattern without word boundaries that was causing match failures
    conditional_oil_pattern = r'(?:an?\s+)?sae\s+(0|5|10|15|20|25)w[-–—]?\s*(\d+)(?:\s+oil)?\s+(?:may\s+be\s+used|can\s+be\s+used|is\s+acceptable)'
    pass_2_5_matches = list(re.finditer(conditional_oil_pattern, text.lower(), re.I))
    for match in pass_2_5_matches:
        base, grade = match.groups()[-2:]
        oil = normalize_oil(f"{base}W-{grade}")
        if oil not in do_not_use_oils:
            context_start = max(0, match.start() - 200)
            context_end = min(len(text), match.end() + 200)
            context = text[context_start:context_end]
            context_lower = context.lower()

            if has_non_engine_oil_context(context_lower) or not has_engine_oil_context(context_lower):
                continue

            if oil not in oil_scores:
                oil_scores[oil] = 0
                oil_temps[oil] = set()
            
            # Lower priority than "best" but still captured
            oil_scores[oil] += 2
            
            # Use the existing extract_temperature function to get temps dynamically
            temps = extract_temperature(context)
            
            # Use extracted temperatures first with a generic fallback when missing.
            final_temps = get_temperature_with_fallback(temps, oil)
            oil_temps[oil].update(final_temps)
            
            # Extract engine specifications from context and map this conditional oil to them
            # This fixes the case where "0W-30 may be used in the 6.0L engine" should link 0W-30 to 6.0L
            rel_match_start = match.start() - context_start
            statement_context = context_lower[rel_match_start:min(len(context_lower), rel_match_start + 200)]

            for eng_m in re.finditer(ENGINE_PATTERN, statement_context):
                eng_size = f"{float(eng_m.group(1)):.1f}L"
                if eng_size not in engine_oil_map_inline:
                    engine_oil_map_inline[eng_size] = []
                if oil not in engine_oil_map_inline[eng_size]:
                    engine_oil_map_inline[eng_size].append(oil)
    
    # PASS 3: Extract from sentences
    sentences = re.split(r'(?<=[.!?])\s+', text)
    for sentence in sentences:
        lower = sentence.lower()
        if re.search(do_not_use_pattern, lower):
            continue
        if has_non_engine_oil_context(lower) or not has_engine_oil_context(lower):
            continue
        oils = re.findall(OIL_PATTERN, sentence, re.I)
        if not oils:
            continue
        
        temps = set()
        if "never goes below" in lower:
            temps.add("above 20°F (-7°C)")
        elif "year-round" in lower or "all temperatures" in lower:
            temps.add("all temperatures")
        else:
            # DYNAMIC: Always use extract_temperature to get actual temp values from the PDF
            temps = extract_temperature(sentence)
        
        for base, grade in oils:
            oil = normalize_oil(f"{base}W-{grade}")
            if oil in do_not_use_oils:
                continue
            if oil not in oil_scores:
                oil_scores[oil] = 0
                oil_temps[oil] = set()
            
            if "preferred" in lower or "recommended" in lower:
                oil_scores[oil] += 5
            elif "may use" in lower or "can use" in lower:
                oil_scores[oil] += 3
            else:
                oil_scores[oil] += 1
            
            # Use extracted temperatures first with a generic fallback when missing.
            final_temps = get_temperature_with_fallback(temps, oil)
            oil_temps[oil].update(final_temps)
    
    # PASS 4: Scan for uncoded oils
    all_oils = re.findall(OIL_PATTERN, text)
    for base, grade in all_oils:
        oil = normalize_oil(f"{base}W-{grade}")
        if oil in do_not_use_oils:
            continue
        if oil not in oil_scores:
            pattern = f"{base}\\s*W[-–—]?\\s*{grade}"
            for match in re.finditer(pattern, text):
                start_pos = max(0, match.start() - 90)
                end_pos = min(len(text), match.end() + 90)
                context = text[start_pos:end_pos].lower()
                if (
                    not has_non_engine_oil_context(context)
                    and has_engine_oil_context(context)
                    and not re.search(do_not_use_pattern, context)
                ):
                    oil_scores[oil] = 1
                    oil_temps[oil] = {"unknown"}
                    break
    
    # Apply document-level temperatures
    doc_temps = extract_temperature(text)
    has_doc_specific_temps = any("F" in t or "°" in t or "range:" in t for t in doc_temps if t != "all temperatures")
    
    if has_doc_specific_temps:
        for oil in oil_scores:
            if "unknown" in oil_temps.get(oil, set()):
                oil_temps[oil] = doc_temps
    elif "never goes below" in text.lower():
        for oil in oil_scores:
            if "unknown" in oil_temps.get(oil, set()):
                oil_temps[oil] = {"above 20°F (-7°C)"}
    elif "year-round" in text.lower():
        for oil in oil_scores:
            if "unknown" in oil_temps.get(oil, set()):
                oil_temps[oil] = {"all temperatures"}
    
    return oil_scores, oil_temps, engine_oil_map_inline

def select_best_engine(engine_caps, all_engines):
    """
    Select the most likely primary engine from candidates.
    
    Step-by-step process:
    1. If engine capacities available: iterate through engine_caps dict
    2. Validate engine matches its capacity expectations
    3. Return first engine with valid capacity match
    4. Fallback: Return first engine from engine_caps with its capacity
    5. If no capacities: rank engines by frequency (most common first)
    6. Filter realistic engines (1.0L to 3.5L range, excluding outliers)
    7. Return most common realistic engine or fallback to most frequent
    
    Args:
        engine_caps: Dict mapping engine sizes to capacity objects
        all_engines: List of all engines found in document
    
    Returns:
        Tuple of (engine_string, capacity_dict) or (None, None)
    """
    if engine_caps:
        for eng, cap in engine_caps.items():
            wf = cap.get("with_filter")
            if wf and engine_matches_capacity(eng, wf):
                return eng, cap
        first_key = next(iter(engine_caps))
        return first_key, engine_caps[first_key]

    if all_engines:
        engine_counts = {}
        for e in all_engines:
            engine_counts[e] = engine_counts.get(e, 0) + 1
        sorted_engines = sorted(engine_counts, key=engine_counts.get, reverse=True)
        
        # Filter realistic engines (1.0L to 3.5L), handling variants like "1.4L Turbo"
        realistic = []
        for e in sorted_engines:
            match = re.search(r'(\d+\.?\d*)', e)
            if match:
                try:
                    disp = float(match.group(1))
                    if 1.0 <= disp <= 3.5:
                        realistic.append(e)
                except (ValueError, AttributeError):
                    pass
        
        best = realistic[0] if realistic else sorted_engines[0]
        return best, None

    return None, None

def extract_all():
    """
    Main orchestrator function for complete PDF extraction pipeline.
    Manages end-to-end data flow from Google Drive to JSON output file.
    
    Step-by-step process:
    1. Authenticate with Google Drive using service account credentials
    2. Discover all PDF files in target FOLDER_ID recursively
    3. For each PDF file:
       a. Extract vehicle info from filename (year-make-model.pdf)
       b. Download PDF from Google Drive
       c. Parse PDF pages and extract full text
       d. Detect vehicle details from PDF content (fallback)
    4. Extract all technical specifications from PDF text:
       - Engine capacities (per-engine if available)
       - Oil types with recommendation scores
       - Oil temperature conditions
       - Engine types (V6, I4, TURBO, etc)
       - Engine-oil proximity mapping
    5. Build per-engine oil recommendation data:
       - Select best engine from candidates
       - Map oils to specific engines via proximity scoring
       - Build recommendation structure per engine
    6. Fallback handling: If no engine data found:
       - Extract generic PDF-level capacity
       - Use all oils with generic/unknown engine label
       - Map document-level temperatures to all oils
    7. Filter and prioritize oil recommendations:
       - Primary oil: highest recommendation score
       - Secondary oils: within -2 points of primary
       - Include any oil with specific temperature conditions
    8. Construct result record with vehicle metadata and engine oil data
    9. Write all vehicle records to JSON output file (structured_results.json)
    10. Display completion message
    
    Args:
        None (uses global constants: FOLDER_ID, OUTPUT_FILE)
    
    Returns:
        None (writes to OUTPUT_FILE)
    """
    service = get_drive_service()
    pdfs = get_all_pdfs(service, FOLDER_ID)
    print(f"\nTotal PDFs found: {len(pdfs)}\n")
    results = {}

    for file in pdfs:
        filename = file["name"]
        print("Processing:", filename)

        year, make, model = parse_filename(filename)
        pdf_stream = download_pdf(service, file["id"])

        with fitz.open(stream=pdf_stream.read(), filetype="pdf") as doc:
            # Analyze PDF type to determine extraction method
            extraction_type, avg_chars = analyze_pdf_type(doc)
            
            # Show extraction method being used with reason
            if extraction_type == "AUTO":
                print(f"  Using text extraction ({avg_chars} chars/page - text-heavy PDF)")
            else:
                print(f"  Using OCR extraction ({avg_chars} chars/page - scanned/manual document)")
            
            if year is None:
                y2, m2, mo2 = detect_vehicle_from_pdf(doc)
                year  = year  or y2
                make  = make  or m2
                model = model or mo2

            # Extract text from all pages
            pages_with_images = []
            text_parts = []
            for page_num, p in enumerate(doc):
                page_text = p.get_text("text")
                text_parts.append(page_text)
                
                # Track pages with images for OCR
                if extraction_type == "MANUAL":
                    images = p.get_images()
                    if images:
                        pages_with_images.append(page_num)
            
            # Keep raw line breaks for table parsers, and cleaned text for broad prose scans.
            raw_text = "\n".join(text_parts)
            text = clean_text(raw_text)
            
            # If this is a manual/scanned PDF with images, run OCR on those pages
            if extraction_type == "MANUAL" and pages_with_images:
                print(f"      Running OCR on {len(pages_with_images)} image page(s)...")
                ocr_text = extract_text_from_images(doc, pages_with_images)
                # Combine OCR text with existing text extraction
                if ocr_text.strip():
                    raw_text = raw_text + "\n" + ocr_text
                    text = clean_text(raw_text)
                    print(f"      OCR extracted {len(ocr_text)} characters")
            
            # PRIORITY 1: Extract engines from specification tables FIRST (most accurate for structured data)
            table_engines = extract_engines_from_spec_table(raw_text)
            if table_engines:
                print(f"      TABLE EXTRACTION found engines: {table_engines}")
            
            # PRIORITY 2: Fall back to general text extraction if no table engines found
            if table_engines:
                all_engines = table_engines
            else:
                all_engines = extract_engines(text)
                if all_engines:
                    print(f"      General extraction found engines: {all_engines}")
            
            all_engines = filter_engine_outliers(all_engines)
            all_engines = consolidate_engine_variants(all_engines)
            
            # PRIORITY 2: Extract engine-capacity pairs from specification tables (validation source)
            engine_caps = extract_engine_capacities(doc)
            
            # CRITICAL: Filter engines by capacity matching
            # Only keep engines that have actual capacity specifications in the PDF
            # This ensures NO hardcoding - only real spec data is used
            if all_engines and engine_caps:
                # Build set of valid engine base sizes from capacity data
                valid_engine_bases = set()
                for cap_eng in engine_caps.keys():
                    valid_engine_bases.add(cap_eng.split()[0])
                
                # Filter all_engines to only those with capacity data
                filtered_engines = []
                for eng in all_engines:
                    eng_base = eng.split()[0]
                    if eng_base in valid_engine_bases:
                        filtered_engines.append(eng)
                
                # Use filtered list if matches found
                if filtered_engines:
                    all_engines = filtered_engines

            # Engine-specific oil-capacity rows are also trusted engine evidence.
            # This catches tables like "4.2L engine oil 6.0 quarts" even when the
            # separate engine detector only saw one of the engine options.
            if engine_caps:
                known_engine_bases = {eng.split()[0] for eng in all_engines}
                for cap_eng in engine_caps.keys():
                    if cap_eng == "unknown_engine":
                        continue
                    cap_base = cap_eng.split()[0]
                    if cap_base not in known_engine_bases:
                        all_engines.append(cap_eng)
                        known_engine_bases.add(cap_base)
            
            oil_scores, oil_temps, engine_oil_map_inline = extract_oils(text)
            
            for oil in oil_temps:
                temps = oil_temps[oil]
                specific_temps = {t for t in temps if "°" in t or "F" in t or "range:" in t or "above" in t or "below" in t}
                if specific_temps and "all temperatures" in temps:
                    oil_temps[oil] = specific_temps
            
            # Extract engine types dynamically from the detected engines only
            engine_types = []
            if all_engines:
                for eng in all_engines:
                    parts = eng.split()
                    for part in parts[1:]:  # Skip displacement, get variants
                        part_upper = part.upper()
                        if part_upper not in engine_types and part_upper not in ["L", "LITER", "LITRE"]:
                            engine_types.append(part_upper)
            
            engine_types_from_text = extract_engine_types(text, all_engines)
            for et in engine_types_from_text:
                if et not in engine_types:
                    engine_types.append(et)
            
            engine_oil_map = map_oils_to_engines(text)
            
            # Merge inline engine-oil mappings (higher priority) with proximity-based mappings
            for eng_size, oils_list in engine_oil_map_inline.items():
                if eng_size not in engine_oil_map:
                    engine_oil_map[eng_size] = oils_list[:]
                else:
                    # Prepend inline oils for higher priority
                    for oil in oils_list:
                        if oil not in engine_oil_map[eng_size]:
                            engine_oil_map[eng_size].insert(0, oil)
            
            # TRUST extract_engine_capacities() completely - it has proper validation
            # Do NOT re-filter or override with all_engines order (causes corruption)
            # Only use fallback if extract_engine_capacities found nothing

            if not engine_caps and all_engines:
                caps = []

                for m in re.finditer(CAPACITY_PATTERN, text):
                    if not is_real_capacity_match(text, m):
                        continue
                    q, l = to_quarts_liters(m.group(1), m.group(2))
                    if 3.0 <= q <= 9.0:
                        caps.append({"quarts": q, "liters": l})

                paired_caps = pair_quarts_liters(caps)

                for i in range(min(len(all_engines), len(paired_caps))):
                    engine_caps[all_engines[i]] = {
                        "with_filter": paired_caps[i],
                        "without_filter": None
                    }

            engine_caps = expand_shared_capacity_to_detected_engines(engine_caps, all_engines)
            engine_caps = align_capacity_engine_keys_with_detected_variants(engine_caps, all_engines)

            selected_engine, selected_cap_entry = select_best_engine(engine_caps, all_engines)
            
            
            multi_engine_data = build_multi_engine_data(engine_caps, oil_scores, oil_temps, engine_oil_map)
            
            # Snapshot the values BEFORE any corrections, so we can validate them
            original_engine_caps = {eng: dict(cap) for eng, cap in engine_caps.items()}
            
            # DISABLED: Corrections are introducing errors by mixing fluid types (oil vs coolant)
            # The initial dynamic extraction is working correctly. Trying to "correct" by finding
            # other capacity mentions in the document is matching wrong contexts (cooling system,
            # transmission, fuel, etc.). Trust the first extraction from PDF.
            # multi_engine_data = apply_correct_capacities(multi_engine_data, doc)
            
            # DYNAMIC NORMALIZATION: Normalize and fix engine data
            if multi_engine_data:
                # 1. Normalize all capacity values (round 6.98 → 7.0)
                for eng_str, eng_info in multi_engine_data.items():
                    cap_info = eng_info.get("oil_capacity", {})
                    with_filter = cap_info.get("with_filter")
                    without_filter = cap_info.get("without_filter")
                    
                    if with_filter and "quarts" in with_filter:
                        with_filter["quarts"] = normalize_capacity_value(with_filter["quarts"])
                    if without_filter and "quarts" in without_filter:
                        without_filter["quarts"] = normalize_capacity_value(without_filter["quarts"])
                
                # 2. Fix unknown_engine by detecting actual engines
                multi_engine_data = fix_unknown_engine(doc, multi_engine_data)

                # 2.5 Dynamic engine-key sanity filter:
                # Keep only engines that also exist in extracted engine candidates.
                # This removes false keys from capacity conversion tokens (e.g., 4.0L, 6.6L).
                if all_engines and multi_engine_data:
                    valid_engine_bases = {eng.split()[0] for eng in all_engines}
                    filtered_multi_engine_data = {
                        eng_key: eng_val
                        for eng_key, eng_val in multi_engine_data.items()
                        if eng_key.split()[0] in valid_engine_bases
                    }
                    if filtered_multi_engine_data:
                        multi_engine_data = filtered_multi_engine_data
                
                # 3. Validate and warn about suspicious capacities
                for eng_str, eng_info in multi_engine_data.items():
                    cap_info = eng_info.get("oil_capacity", {})
                    with_filter = cap_info.get("with_filter")
                    if with_filter and "quarts" in with_filter:
                        cap_q = with_filter["quarts"]
                        if not (1.0 <= cap_q <= 12.0):
                            print(f"      Warning: {eng_str} capacity {cap_q}qt seems outside normal automotive oil range")
            
            if not multi_engine_data:
                fallback_cap = extract_fallback_capacity(doc)

                oil_list = []
                if oil_scores:
                    primary = max(oil_scores, key=oil_scores.get)
                    max_score = oil_scores[primary]
                    
                    for oil, score in oil_scores.items():
                        temps = oil_temps.get(oil, [])
                        has_actual_temps = any(
                            ("°F" in t or "°C" in t or "F" in t or "C" in t or 
                             "above" in t or "below" in t or "range:" in t or 
                             "weather" in t or "temperatures" in t or "cold" in t or "hot" in t)
                            for t in temps
                        )
                        
                        if score >= max_score - 2 or has_actual_temps:
                            oil_list.append({
                                "oil_type": oil,
                                "recommendation_level": "primary" if oil == primary else "secondary",
                                "temperature_condition": list(temps)
                            })

                if fallback_cap:
                    if all_engines:
                        multi_engine_data = {
                            eng: {
                                "oil_capacity": {
                                    "with_filter": dict(fallback_cap["with_filter"]) if fallback_cap.get("with_filter") else None,
                                    "without_filter": dict(fallback_cap["without_filter"]) if fallback_cap.get("without_filter") else None
                                },
                                "oil_recommendations": oil_list
                            }
                            for eng in all_engines
                        }
                    else:
                        multi_engine_data = {
                            "unknown_engine": {
                                "oil_capacity": fallback_cap,
                                "oil_recommendations": oil_list
                            }
                        }
            
            with_filter_cap    = selected_cap_entry.get("with_filter")    if selected_cap_entry else None
            without_filter_cap = selected_cap_entry.get("without_filter") if selected_cap_entry else None

            oil_list = []
            if oil_scores:
                primary   = max(oil_scores, key=oil_scores.get)
                max_score = oil_scores[primary]
                for oil, score in oil_scores.items():
                    if score >= max_score - 2 or any(
                        "weather" in t or "temperatures" in t
                        for t in oil_temps[oil]
                    ):
                        oil_list.append({
                            "oil_type": oil,
                            "recommendation_level": "primary" if oil == primary else "secondary",
                            "temperature_condition": list(oil_temps[oil]),
                        })
            
            results[filename] = {
                "Vehicle": {
                    "year": year,
                    "make": make,
                    "model": model,
                    "engine_types": list(engine_types),
                    "displayName": f"{year} {make} {model}"
                },
                "engines": multi_engine_data
            }

    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(results, f, indent=4, ensure_ascii=False)

    print("\nExtraction Complete\n")


if __name__ == "__main__":
    extract_all()
