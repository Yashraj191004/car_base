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
CREDS_PATH = "credentials.json"
OUTPUT_FILE = "structured_results.json"
REFERENCE_FILE = "vehicle_reference.json"


# Core text patterns. These stay generic: they describe syntax commonly found in
# manuals, not any specific vehicle. Unicode escapes allow PDF dash/minus/degree
# variants without putting corrupted characters in the source file.
OIL_PATTERN = r"\b(0|5|10|15|20|25)\s*W[-\u2013\u2014]?\s*(16|20|30|40|50|60)\b"
CAPACITY_PATTERN = r"(\d+\.?\d*)\s*(?:us\s+|imp\s+|u\.s\.\s+)?(quarts?|qts?|qt\.?|liters?|l\b)"
CAPACITY_UNIT_PATTERN = r"(?:quarts?|qts?|qt\.?|liters?|litres?|l\b|gal|gallons?|ml|cc)"
ENGINE_PATTERN = r"\b(\d{1,2}\.\d)\s*(?:[-]?\s*(?:l|liter|litre)|(?:\s+(?:gdi|dohc|sohc|turbo|ecoboost|naturally|cylinder)))\b"
ENGINE_TYPE_PATTERN = r"\b(?:v\s*-?\s*(?:3|4|5|6|8|10|12|16|20|24)|i\s*-?\s*[3-8]|l\s*-?\s*[3-8]|w\s*-?\s*(?:8|12|16)|h\s*-?\s*4|f\s*-?\s*8|inline\s*-?\s*[3-8]|flat\s*-?\s*(?:3|4|6|8|12)|boxer|rotary|wankel|turbo(?:charged)?|supercharged|naturally\s*-?\s*aspirated|hybrid|electric|diesel|petrol|gdi|sohc|dohc)\b"
TEMP_PATTERN = r"([\u2212-]?\d+)\s*(?:\u00b0|\u00ba|\u00c2\u00b0)?\s*(c|f)"

ENGINE_MIN_DISPLACEMENT = 0.6
ENGINE_MAX_DISPLACEMENT = 12.0
ENGINE_OIL_CAPACITY_MIN_QT = 1.0
ENGINE_OIL_CAPACITY_MAX_QT = 20.0

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


def get_engine_displacement(engine_text):
    """Return the numeric displacement from an engine label, or None."""
    match = re.search(r'(\d+\.?\d*)', str(engine_text))
    if not match:
        return None

    try:
        return float(match.group(1))
    except (ValueError, TypeError):
        return None


def is_plausible_engine_displacement(value):
    """Keep displacement limits broad; context filters remove most false positives."""
    return value is not None and ENGINE_MIN_DISPLACEMENT <= value <= ENGINE_MAX_DISPLACEMENT


def normalize_engine_type_token(token):
    """Normalize engine type text extracted near a displacement."""
    if not token:
        return ""

    normalized = re.sub(r"\s+", "", token.strip().upper())
    normalized = normalized.replace("INLINE", "I")

    if re.fullmatch(r"[VIWLHF]-?\d{1,2}", normalized):
        normalized = normalized.replace("-", "")

    if re.fullmatch(r"L[3-8]", normalized):
        return "I" + normalized[1:]
    if normalized in ("BOXE", "BOXER"):
        return "BOXER"
    if re.fullmatch(r"FLAT-?\d+", normalized):
        return "FLAT-" + re.search(r"\d+", normalized).group()
    if normalized in ("TWINTTURBO", "TWIN-TURBO"):
        return "TWIN-TURBO"
    if normalized == "BITURBO":
        return "BI-TURBO"
    if normalized == "TURBOCHARGED":
        return "TURBO"

    return normalized


def extract_engine_variant_from_context(context, base_engine=""):
    """Find an engine layout/variant token near a displacement."""
    if not context:
        return ""

    base_upper = base_engine.upper()
    for match in re.finditer(ENGINE_TYPE_PATTERN, context, re.I):
        token = normalize_engine_type_token(match.group())
        if not token or token in base_upper:
            continue
        if token in {"GDI", "SOHC", "DOHC", "DIESEL", "PETROL", "HYBRID", "ELECTRIC"}:
            continue
        return f" {token.title()}" if token in {"TURBO", "SUPERCHARGED"} else f" {token}"

    return ""


def has_engine_signal(text):
    """Return True when a line has enough context to be treated as engine-related text."""
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
    """Detect engine-pattern matches that are actually liter capacity conversions."""
    if not text or not match:
        return False

    context_before = text[max(0, match.start() - 40):match.start()]
    return bool(
        re.search(r"\d+\.?\d*\s*" + CAPACITY_UNIT_PATTERN + r"\s*\(?\s*$", context_before, re.I)
        or re.search(r"\(\s*$", context_before, re.I)
    )


def is_parenthesized_capacity_conversion(text, match):
    """Detect parenthesized metric conversions such as "6.0 qt (5.7L)"."""
    if not text or not match:
        return False

    before = text[max(0, match.start() - 35):match.start()]
    after = text[match.end():min(len(text), match.end() + 5)]
    return bool(
        re.search(r"\d+\.?\d*\s*(?:quarts?|qts?|qt\.?|gal|gallons?)\s*\(\s*$", before, re.I)
        and re.search(r"^\s*\)", after)
    )


def is_capacity_or_fluid_row(text):
    """Identify fluid/capacity rows so table parsers do not treat them as engine spec rows."""
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
    """Separate real fluid capacities from bare engine displacements that also use liters."""
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


def overlaps_real_capacity_match(text, engine_match, capacity_matches=None):
    """Return True when an engine-sized token overlaps a real capacity match."""
    if not text or not engine_match:
        return False

    matches = capacity_matches
    if matches is None:
        matches = [
            cm for cm in re.finditer(CAPACITY_PATTERN, text, re.I)
            if is_real_capacity_match(text, cm)
        ]

    for cap_match in matches:
        if not is_real_capacity_match(text, cap_match):
            continue
        if engine_match.start() < cap_match.end() and cap_match.start() < engine_match.end():
            return True

    return False


def get_temperature_with_fallback(extracted_temps, oil_type):
    """
    Normalize extracted temperature labels; default to all temperatures when the PDF
    gives no condition.
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
    """Return True when nearby wording shows an oil mention belongs to another fluid system."""
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
    """Require positive engine-oil wording before accepting an oil recommendation."""
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
        "viscosity grade",
        "best viscosity grade",
        "cold temperature operation",
        "selecting an oil",
        "correct specification",
        "year-round protection",
        "improved fuel economy",
        "oil container",
    ]
    return any(signal in text_lower for signal in engine_oil_signals)



def extract_text_from_images(doc, pages_with_images):
    """OCR image-heavy pages and return their combined text."""
    ocr_text = []
    
    for page_num in pages_with_images:
        try:
            page = doc[page_num]
            pix = page.get_pixmap(matrix=fitz.Matrix(4, 4))
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            text = pytesseract.image_to_string(img)
            if text.strip():
                ocr_text.append(text)
        except Exception as e:
            print(f"      Warning: OCR failed on page {page_num}: {str(e)}")
            continue
    
    return " ".join(ocr_text)


def get_drive_service():
    """Create an authenticated Google Drive client from the service account credentials."""
    creds = service_account.Credentials.from_service_account_file(
        CREDS_PATH, scopes=SCOPES
    )
    return build("drive", "v3", credentials=creds)


def get_all_pdfs(service, folder_id):
    """Recursively collect PDF files under the configured Google Drive folder."""
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
    """Use average extracted text per page to choose direct text extraction or OCR assist."""
    total_chars = 0
    page_count = len(doc)
    
    for page in doc:
        text = page.get_text("text")
        total_chars += len(text)
    
    avg_chars = int(total_chars / page_count) if page_count > 0 else 0
    
    extraction_type = "AUTO" if avg_chars >= 800 else "MANUAL"
    
    return extraction_type, avg_chars


def download_pdf(service, file_id):
    """Download a Google Drive PDF into an in-memory byte buffer."""
    request = service.files().get_media(fileId=file_id)
    buffer = io.BytesIO()
    downloader = MediaIoBaseDownload(buffer, request)
    done = False
    
    while not done:
        _, done = downloader.next_chunk()
    
    buffer.seek(0)
    return buffer


def clean_text(text):
    """Collapse whitespace for broad prose matching while table parsers keep raw line breaks."""
    return re.sub(r"\s+", " ", text.replace("\n", " "))


def normalize_oil(raw):
    """Normalize oil viscosity text to the canonical form, for example 5W-30."""
    oil = raw.upper().replace("\u2013", "-").replace("\u2014", "-")
    oil = re.sub(r"\s+", "", oil)
    if "W-" not in oil:
        oil = oil.replace("W", "W-")
    return oil


def to_quarts_liters(value, unit):
    """Convert one capacity value to both quarts and liters."""
    value = float(value)
    if unit.lower().startswith("l"):
        return round(value / 0.946, 2), value
    else:
        return value, round(value * 0.946, 1)



def parse_filename(name):
    """Extract year, make, and model from filenames such as 2017-Honda-Civic-OM.pdf."""
    clean_name = name.replace("Copy of ", "").strip()
    
    match = re.match(r"(\d{4})-([^-]+)-(.+)\.pdf", clean_name, re.I)
    if match:
        year, make, model = match.groups()
        model = model.replace("-OM", "").replace("-UG", "").replace("-UM", "")
        return int(year), make.capitalize(), model.capitalize()
    return None, None, None


def build_multi_engine_data(engine_caps, oil_scores, oil_temps, engine_oil_map):
    """Build the final per-engine capacity and oil recommendation records."""
    engine_data = {}

    for eng, cap in engine_caps.items():
        with_filter = cap.get("with_filter")
        without_filter = cap.get("without_filter")
        oil_list = []

        if oil_scores:
            valid_oils = engine_oil_map.get(eng, [])
            
            if not valid_oils:
                base_eng_size = eng.split()[0]
                valid_oils = engine_oil_map.get(base_eng_size, [])
            
            if not valid_oils:
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
                    ("F" in t or "C" in t or 
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
    Copy a shared unknown-engine capacity to detected engines when the PDF gives no
    per-engine rows.
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
    """Use richer detected engine keys, such as 1.4L Turbo, for matching capacity rows."""
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


def filter_engine_caps_to_detected_engines(engine_caps, all_engines):
    """Drop capacity keys that do not match trusted detected engine bases."""
    if not engine_caps or not all_engines:
        return engine_caps

    valid_bases = {eng.split()[0] for eng in all_engines}
    filtered = {}
    unknown_cap = engine_caps.get("unknown_engine")

    for eng, cap in engine_caps.items():
        if eng == "unknown_engine":
            continue

        if eng.split()[0] in valid_bases:
            filtered[eng] = cap

    if filtered:
        return filtered
    if unknown_cap:
        return {"unknown_engine": unknown_cap}
    return engine_caps


def prefer_shared_capacity_if_current_caps_are_noise(engine_caps, shared_cap):
    """
    Use a clear shared oil capacity when current engine-specific caps look like
    stray part-table values. This protects rows after headers like "Oil Filter".
    """
    if not engine_caps or not shared_cap or "unknown_engine" in engine_caps:
        return engine_caps

    shared_wf = shared_cap.get("with_filter") or {}
    shared_q = shared_wf.get("quarts")
    if not shared_q:
        return engine_caps

    current_qs = []
    for cap in engine_caps.values():
        wf = (cap or {}).get("with_filter") or {}
        q = wf.get("quarts")
        if isinstance(q, (int, float)):
            current_qs.append(q)

    if current_qs and max(current_qs) <= 2.0 and shared_q > max(current_qs) + 1.0:
        return {"unknown_engine": shared_cap}

    return engine_caps


def apply_shared_capacity_to_noisy_engine_data(engine_data, shared_cap):
    """
    Repair final engine records when a clear shared oil capacity exists but
    every detected engine ended up with tiny values from nearby non-capacity rows.
    """
    if not engine_data or not shared_cap:
        return engine_data

    shared_wf = shared_cap.get("with_filter") or {}
    shared_q = shared_wf.get("quarts")
    if not isinstance(shared_q, (int, float)):
        return engine_data

    current_qs = []
    for info in engine_data.values():
        cap_info = (info or {}).get("oil_capacity") or {}
        wf = cap_info.get("with_filter") or {}
        q = wf.get("quarts")
        if isinstance(q, (int, float)):
            current_qs.append(q)

    if not current_qs or max(current_qs) > 2.0 or shared_q <= max(current_qs) + 1.0:
        return engine_data

    for info in engine_data.values():
        cap_info = info.setdefault("oil_capacity", {})
        cap_info["with_filter"] = {k: v for k, v in shared_wf.items() if k != "pos"}
        if "without_filter" not in cap_info:
            cap_info["without_filter"] = shared_cap.get("without_filter")

    return engine_data


def engine_label_has_type(engine_label):
    """Return True when an engine key already includes a layout or variant token."""
    if not engine_label or engine_label == "unknown_engine":
        return True

    parts = str(engine_label).split()
    if len(parts) <= 1:
        return False

    trailing = " ".join(parts[1:])
    return bool(re.search(ENGINE_TYPE_PATTERN, trailing, re.I))


def select_single_layout_type(engine_types):
    """Pick one vehicle-wide layout type only when it is unambiguous."""
    layout_types = []
    for engine_type in engine_types or []:
        normalized = normalize_engine_type_token(engine_type)
        if re.fullmatch(r"(?:I|V|W|H|F)\d{1,2}", normalized) or normalized.startswith("FLAT") or normalized in {"BOXER", "ROTARY", "WANKEL"}:
            if normalized not in layout_types:
                layout_types.append(normalized)

    return layout_types[0] if len(layout_types) == 1 else None


def add_missing_engine_type_to_keys(engine_data, engine_types):
    """
    Append an unambiguous vehicle-wide layout to bare displacement keys.
    Example: {"1.4L": ...} with engine_types ["I4"] becomes {"1.4L I4": ...}.
    """
    layout_type = select_single_layout_type(engine_types)
    if not engine_data or not layout_type:
        return engine_data

    relabeled = {}
    for engine_label, engine_info in engine_data.items():
        if engine_label_has_type(engine_label):
            relabeled[engine_label] = engine_info
            continue

        relabeled[f"{engine_label} {layout_type}"] = engine_info

    return relabeled


def pair_quarts_liters(cap_list):
    """Pair sequential quart/liter capacity values from fallback extraction."""
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
    """Load optional make/model reference data used only for fallback repair."""
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
    """Normalize make/model text for case-insensitive reference matching."""
    if not value:
        return ""

    text = value.lower()
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def simplify_model_for_lookup(model):
    """Remove body-style words so a detected model can match the reference list."""
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
    """Identify low-confidence make text that should be eligible for fallback repair."""
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
    """Check a detected make against the optional reference list."""
    if not make:
        return False

    if reference is None:
        reference = load_vehicle_reference()

    make_key = normalize_vehicle_label(make)
    known_keys = {normalize_vehicle_label(ref_make) for ref_make in reference.keys()}
    return make_key in known_keys


def resolve_make_from_model_reference(model, detected_make=None):
    """Use the reference list to repair a missing or suspicious make from the detected model."""
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
    """Infer year, make, and model from the first pages when filename parsing is unavailable."""
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
    """Associate oils with nearby engine mentions using local text proximity."""
    engine_oil_map = {}
    engine_positions = []
    
    for m in re.finditer(ENGINE_PATTERN, text):
        displacement = float(m.group(1))
        if not is_plausible_engine_displacement(displacement):
            continue
        eng = f"{displacement:.1f}L"
        engine_positions.append((eng, m.start()))

    for eng, eng_pos in engine_positions:
        window = text[max(0, eng_pos - 200): eng_pos + 300]
        oils = re.findall(OIL_PATTERN, window)
        engine_oil_map[eng] = list(set([
            normalize_oil(f"{b}W-{g}") for b, g in oils
        ]))

    return engine_oil_map


def extract_engines(text):
    """Extract engine displacements and variants from prose while avoiding capacity rows."""
    engines = []
    text_lower = text.lower()
    
    skip_phrases = [
        "previous generation", "previous model", "prior generation",  
        "towing capacity", "cargo capacity", "payload capacity",
        "weight rating", "gvwr", "gcwr", "optional", "as an option",
        "engine oil", "transaxle fluid", "cooling system", "fuel tank"
    ]
    
    sentences = re.split(r'[.!?]\s+', text)
    
    for sentence in sentences:
        sentence_lower = sentence.lower()
        
        if any(phrase in sentence_lower for phrase in skip_phrases):
            continue
        
        if not has_engine_signal(sentence_lower):
            continue
        
        if any(ctx in sentence_lower for ctx in NON_ENGINE_CONTEXT):
            continue
        
        for m in re.finditer(ENGINE_PATTERN, sentence_lower):
            try:
                num = float(m.group(1).strip())
                if is_plausible_engine_displacement(num):
                    eng_str = f"{num:.1f}L"
                    
                    # Capacity tables often contain values that look like engines, e.g.
                    # "4.0 L 4.2 qt". Reject matches sitting in capacity context.
                    context_before = sentence_lower[max(0, m.start() - 30):m.start()]
                    capacity_indicators = ['qt', 'quart', 'liter', 'litre', 'gal', 'gallon', 'ml', 'cc']
                    if any(cap_ind in context_before for cap_ind in capacity_indicators) or is_capacity_conversion_engine_match(sentence_lower, m) or is_parenthesized_capacity_conversion(sentence_lower, m):
                        continue
                    
                    window_start = max(0, m.start() - 50)
                    window_end = min(len(sentence_lower), m.end() + 100)
                    window = sentence_lower[window_start:window_end]
                    variant = extract_engine_variant_from_context(window, eng_str)
                    
                    full_eng = eng_str + variant
                    if full_eng not in engines:
                        engines.append(full_eng)
            except (ValueError, TypeError):
                continue
    
    # No static code-to-displacement fallback. Keep only engines found in document text.
    return list(set(engines))


def extract_engines_from_spec_table(text):
    """Extract engines from structured specification tables before falling back to prose."""
    table_engines = []
    lines = text.split('\n')
    
    # First find likely engine-spec table headers, then inspect nearby rows.
    spec_header_indices = []
    for idx, line in enumerate(lines):
        line_lower = line.lower().strip()
        
        if is_capacity_or_fluid_row(line_lower):
            continue

        # Keep this strict so "Engine Oil with Filter" is not treated as a spec table.
        has_engine_spec_title = "engine" in line_lower and ("spec" in line_lower or "data" in line_lower)
        has_engine_table_header = "engine" in line_lower and any(
            kw in line_lower for kw in ['vin', 'transaxle', 'spark', 'code', 'gap']
        )
        
        has_table_keywords = any(kw in line_lower for kw in ['vin', 'transaxle', 'spark', 'code', 'gap', 'cubic inches', 'firing order', 'compression ratio'])
        
        if has_engine_spec_title or has_engine_table_header or (has_table_keywords and len(spec_header_indices) < 10):
            spec_header_indices.append(idx)
    
    for header_idx in spec_header_indices:
        for row_idx in range(header_idx + 1, min(header_idx + 20, len(lines))):
            row = lines[row_idx].strip()
            
            if not row or len(row) < 3:
                continue
            
            if any(sep in row for sep in ['---', '===', '***']):
                continue
            
            row_lower = row.lower()

            if is_capacity_or_fluid_row(row_lower):
                continue
            
            has_multi_engine_header = row_lower.startswith("engine ") and len(list(re.finditer(ENGINE_PATTERN, row_lower))) >= 2
            if not has_engine_signal(row_lower) and not has_multi_engine_header:
                continue
            
            if any(kw in row_lower for kw in ['cargo', 'towing', 'payload', 'fuel tank', 'weight', 'gvwr']):
                continue
            
            engines_in_row = []
            row_engine_matches = list(re.finditer(ENGINE_PATTERN, row_lower))
            for match_idx, eng_match in enumerate(row_engine_matches):
                try:
                    # Reject capacity conversions before accepting a table cell as an engine.
                    context_before = row_lower[max(0, eng_match.start() - 30):eng_match.start()]
                    capacity_indicators = ['qt', 'quart', 'liter', 'litre', 'gal', 'gallon', 'ml', 'cc']
                    if any(cap_ind in context_before for cap_ind in capacity_indicators) or is_capacity_conversion_engine_match(row_lower, eng_match) or is_parenthesized_capacity_conversion(row_lower, eng_match):
                        continue
                    
                    eng_size_num = float(eng_match.group(1))
                    if is_plausible_engine_displacement(eng_size_num):
                        eng_str = f"{eng_size_num:.1f}L"
                        
                        # Look for engine variant in this engine's cell/segment, not the next engine column.
                        next_engine_start = row_engine_matches[match_idx + 1].start() if match_idx + 1 < len(row_engine_matches) else len(row_lower)
                        row_window = row_lower[max(0, eng_match.start() - 15):next_engine_start]
                        variant = extract_engine_variant_from_context(row_window, eng_str)
                        
                        full_eng = eng_str + variant
                        if full_eng not in engines_in_row:
                            engines_in_row.append(full_eng)
                except (ValueError, TypeError):
                    continue
            
            table_engines.extend(engines_in_row)
    
    return list(set(table_engines))


def filter_engine_outliers(engines):
    """Remove displacement outliers that are likely unrelated capacity or reference values."""
    if not engines:
        return engines
    
    if len(engines) <= 2:
        return engines
    
    engine_nums = []
    for eng in engines:
        try:
            disp = get_engine_displacement(eng)
            if is_plausible_engine_displacement(disp):
                engine_nums.append((disp, eng))
        except (ValueError, AttributeError):
            continue
    
    if not engine_nums:
        return engines
    
    engine_nums.sort()
    
    # Four or more engines usually means capacity/reference noise. Split on a
    # large displacement gap and keep the more plausible cluster.
    if len(engine_nums) >= 4:
        gaps = []
        for i in range(len(engine_nums) - 1):
            gap = engine_nums[i + 1][0] - engine_nums[i][0]
            gaps.append((gap, i))
        
        if gaps:
            for gap_size, gap_idx in sorted(gaps, reverse=True):
                if gap_size > 1.2:
                    lower_group = engine_nums[:gap_idx + 1]
                    upper_group = engine_nums[gap_idx + 1:]
                    
                    lower_has_large = any(d > 5.5 for d, _ in lower_group)
                    upper_has_large = any(d > 5.5 for d, _ in upper_group)
                    lower_has_small = any(d < 1.8 for d, _ in lower_group)
                    upper_has_small = any(d < 1.8 for d, _ in upper_group)
                    lower_has_explicit_type = any(re.search(ENGINE_TYPE_PATTERN, e, re.I) for _, e in lower_group)
                    upper_has_explicit_type = any(re.search(ENGINE_TYPE_PATTERN, e, re.I) for _, e in upper_group)
                    
                    if lower_has_small and not upper_has_small and upper_has_large and not upper_has_explicit_type:
                        engine_nums = lower_group
                        break
                    elif upper_has_small and not lower_has_small and lower_has_large and not lower_has_explicit_type:
                        engine_nums = upper_group
                        break
                    elif upper_has_explicit_type and not lower_has_explicit_type:
                        engine_nums = upper_group
                        break
                    elif lower_has_explicit_type and not upper_has_explicit_type:
                        engine_nums = lower_group
                        break
                    elif len(lower_group) >= len(upper_group):
                        engine_nums = lower_group
                        break
                    else:
                        engine_nums = upper_group
                        break
    
    # Keep one best label per displacement, preferring typed variants.
    displacement_map = {}
    for disp, eng_str in engine_nums:
        base_disp = round(disp, 1)
        
        if base_disp not in displacement_map:
            displacement_map[base_disp] = eng_str
        else:
            current = displacement_map[base_disp]
            current_has_type = bool(re.search(ENGINE_TYPE_PATTERN, current, re.I))
            new_has_type = bool(re.search(ENGINE_TYPE_PATTERN, eng_str, re.I))
            
            if new_has_type and not current_has_type:
                displacement_map[base_disp] = eng_str
            elif new_has_type and current_has_type:
                if len(eng_str) > len(current):
                    displacement_map[base_disp] = eng_str
    
    result = list(displacement_map.values())[:6]
    
    # 3.8L is often a fuel/capacity false positive when a larger typed cluster exists.
    if result and any(eng for eng in result if '3.8' in str(eng)):
        has_larger_typed_engine = any(
            (get_engine_displacement(eng) or 0) >= 5.0 and re.search(ENGINE_TYPE_PATTERN, eng, re.I)
            for eng in result
        )
        if has_larger_typed_engine:
            result = [eng for eng in result if '3.8' not in str(eng)]
    
    return result if result else engines


def consolidate_engine_variants(engines):
    """Prefer typed variants over duplicate base engine keys for the same displacement."""
    if not engines:
        return engines
    
    groups = {}
    for eng in engines:
        match = re.search(r'(\d+\.?\d*L)', eng)
        if match:
            base = match.group(1)
            if base not in groups:
                groups[base] = []
            groups[base].append(eng)
    
    result = []
    for base, eng_list in groups.items():
        if len(eng_list) > 1:
            variants = [e for e in eng_list if e != base]
            if variants:
                result.extend(variants)
            else:
                result.extend(eng_list)
        else:
            result.extend(eng_list)
    
    return result


def has_engine_context(text, match_pos, match_text, context_window=150):
    """Validate that an engine-type token appears near actual engine wording."""
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
    Extract engine layouts and variants from detected engine rows, with
    displacement-based fallback.
    """
    engine_types_found = set()
    text_lower = text.lower()
    lines = text_lower.split('\n')
    
    # Pass 1: read explicit type tokens attached to engine-size rows.
    for line in lines:
        capacity_matches = [
            cm for cm in re.finditer(CAPACITY_PATTERN, line, re.I)
            if is_real_capacity_match(line, cm)
        ]
        engine_matches = list(re.finditer(ENGINE_PATTERN, line))
        if not engine_matches:
            continue
        
        for engine_match in engine_matches:
            displacement = float(engine_match.group(1))
            if (
                not is_plausible_engine_displacement(displacement)
                or overlaps_real_capacity_match(line, engine_match, capacity_matches)
            ):
                continue

            engine_start = engine_match.start()
            engine_end = engine_match.end()
            engine_size_text = engine_match.group()
            
            end_search = min(len(line), engine_end + 50)
            trailing_text = line[engine_end:end_search]
            
            delimiter_pos = len(trailing_text)
            for delim in [',', ';', '(', ')']:
                pos = trailing_text.find(delim)
                if pos != -1 and pos < delimiter_pos:
                    delimiter_pos = pos
            
            descriptor_text = (engine_size_text + trailing_text[:delimiter_pos]).lower()
            
            mechanical_pattern = r"\b(v3|v4|v5|v6|v8|v10|v12|v16|v20|v24|i3|i4|i5|i6|i8|l3|l4|l5|l6|l8|h4|w8|w12|w16|f8|flat[-]?4|flat[-]?6|flat[-]?12|boxer|boxe|rotary|wankel|turbo|dual[-]?turbo|quad[-]?turbo|sequential[-]?turbo|supercharged|twin[-]?supercharged|twin[-]?turbo|twin[-]?scroll)\b"
            
            for match in re.finditer(mechanical_pattern, descriptor_text):
                match_text = normalize_engine_type_token(match.group())
                engine_types_found.add(match_text)
            
            inline_pattern = r"\binline\s*[-]?\s*([3-8])\b"
            for match in re.finditer(inline_pattern, descriptor_text):
                digit = match.group(1)
                engine_types_found.add(f"I{digit}")
            
            cylinder_pattern = r"\b([3-8])\s*[-]?\s*cylinder\b"
            for match in re.finditer(cylinder_pattern, descriptor_text):
                digit = match.group(1)
                engine_types_found.add(f"I{digit}")
    
    # Pass 2: if the PDF does not spell out a type, infer from detected engines
    # only. Do not use every displacement mention in the PDF.
    engine_displacements = []
    if all_engines:
        for eng in all_engines:
            displacement = get_engine_displacement(eng)
            if is_plausible_engine_displacement(displacement):
                engine_displacements.append(displacement)
    
    if engine_displacements:
        min_disp = min(engine_displacements)
        max_disp = max(engine_displacements)
        
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
            # Reject obvious mismatches, such as V8 near only sub-3.0L engines.
            invalid_types = set()
            for et in engine_types_found:
                if et in ("V8", "V10", "V12", "V16") and max_disp < 3.0:
                    invalid_types.add(et)
                elif et in ("I3", "I4", "I5", "I6") and min_disp > 4.0:
                    invalid_types.add(et)
            
            if invalid_types:
                engine_types_found -= invalid_types
                
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
    
    normalized_types = set()
    for engine_type in engine_types_found:
        normalized = normalize_engine_type_token(engine_type)
        if normalized:
            normalized_types.add(normalized)
    
    return sorted(list(normalized_types))


def engine_matches_capacity(engine, capacity):
    """Check whether an engine displacement and oil capacity are plausibly paired."""
    if not capacity:
        return True
    
    quarts = capacity.get("quarts", 0)
    if not (ENGINE_OIL_CAPACITY_MIN_QT <= quarts <= ENGINE_OIL_CAPACITY_MAX_QT):
        return False
    
    size = get_engine_displacement(engine)
    if size is None:
        return True

    if not is_plausible_engine_displacement(size):
        return False

    if size <= 2.0 and quarts <= 8.0:
        return True
    if size <= 4.0 and quarts <= 12.0:
        return True
    if size > 4.0 and quarts >= 4.0:
        return True

    return False


def extract_temperature(sentence):
    """Extract temperatures, convert Celsius to Fahrenheit, and classify weather conditions."""
    s = sentence.lower()
    temps = re.findall(TEMP_PATTERN, s)
    values = []
    
    for value, unit in temps:
        value = int(value.replace("\u2212", "-"))
        if unit == "c":
            value = round((value * 9 / 5) + 32)
        if -60 <= value <= 150:
            values.append(value)
    
    values = sorted(set(values))
    result = set()
    
    if not values:
        return {"all temperatures"}
    
    min_temp, max_temp = min(values), max(values)
    
    has_cold = any(t <= 40 for t in values)
    has_hot = any(t >= 85 for t in values)
    
    if has_cold and has_hot:
        result.add("all temperatures")
        if not (min_temp <= -50 and max_temp >= 140):
            result.add(f"range: {min_temp}F to {max_temp}F")
    elif has_cold:
        result.add("cold weather")
        if len(values) > 1:
            result.add(f"range: {min_temp}F to {max_temp}F")
        else:
            result.add(f"{values[0]}F")
    elif has_hot:
        result.add("hot weather")
        if len(values) > 1:
            result.add(f"range: {min_temp}F to {max_temp}F")
        else:
            result.add(f"{values[0]}F")
    else:
        if len(values) <= 3:
            for temp in values:
                result.add(f"{temp}F")
        else:
            result.add(f"range: {min_temp}F to {max_temp}F")
    
    return result


def normalize_capacity_value(quarts_value):
    """Round capacity values only when they are very close to a whole number."""
    if not isinstance(quarts_value, (int, float)):
        return quarts_value
    
    rounded = round(quarts_value)
    if abs(quarts_value - rounded) < 0.05:
        return float(rounded)
    return round(quarts_value, 2)


def find_correct_engine_oil_capacities(doc):
    """Legacy helper that searches explicit engine-oil lines for capacity corrections."""
    correct_capacities = {}
    
    for page_num, page in enumerate(doc):
        text = page.get_text()
        text_lower = text.lower()
        
        if not ("capacity" in text_lower or "specification" in text_lower):
            continue
        
        lines = text.split('\n')
        
        for line_idx, line in enumerate(lines):
            line_lower = line.lower()
            
            if "engine oil" not in line_lower:
                continue
            
            if "coolant" in line_lower or "transmission" in line_lower or "cooling" in line_lower:
                continue
            
            engines = [
                (float(m.group(1)), m.start())
                for m in re.finditer(ENGINE_PATTERN, line_lower)
                if is_plausible_engine_displacement(float(m.group(1)))
            ]
            capacities = [(m.group(1), m.group(2), m.start()) for m in re.finditer(CAPACITY_PATTERN, line_lower, re.IGNORECASE)]
            
            if engines and capacities:
                for eng_val, eng_pos in engines:
                    eng_size = f"{eng_val:.1f}L"
                    
                    for cap_str, cap_unit, cap_pos in capacities:
                        if cap_pos > eng_pos:
                            try:
                                q, l = to_quarts_liters(cap_str, cap_unit)
                                if ENGINE_OIL_CAPACITY_MIN_QT <= q <= ENGINE_OIL_CAPACITY_MAX_QT:
                                    correct_capacities[eng_size] = {
                                        "quarts": round(q, 2),
                                        "liters": round(l, 1)
                                    }
                                    break
                            except (ValueError, TypeError):
                                continue
            else:
                if engines and line_idx + 1 < len(lines):
                    next_line = lines[line_idx + 1]
                    next_lower = next_line.lower()
                    next_capacities = [(m.group(1), m.group(2), m.start()) for m in re.finditer(CAPACITY_PATTERN, next_lower, re.IGNORECASE)]
                    
                    if next_capacities and "engine oil" not in next_lower:
                        for eng_val, _ in engines:
                            eng_size = f"{eng_val:.1f}L"
                            
                            for cap_str, cap_unit, _ in next_capacities:
                                try:
                                    q, l = to_quarts_liters(cap_str, cap_unit)
                                    if ENGINE_OIL_CAPACITY_MIN_QT <= q <= ENGINE_OIL_CAPACITY_MAX_QT:
                                        correct_capacities[eng_size] = {
                                            "quarts": round(q, 2),
                                            "liters": round(l, 1)
                                        }
                                        break
                                except (ValueError, TypeError):
                                    continue
    
    return correct_capacities


def apply_correct_capacities(multi_engine_data, doc):
    """Legacy correction pass that compares current capacities against other oil-only mentions."""
    if not multi_engine_data:
        return multi_engine_data
    
    for eng_str in list(multi_engine_data.keys()):
        cap_info = multi_engine_data[eng_str].get("oil_capacity", {})
        with_filter = cap_info.get("with_filter")
        
        if not with_filter or "quarts" not in with_filter:
            continue
        
        current_q = with_filter["quarts"]
        
        all_capacities = extract_all_capacities_for_engine(doc, eng_str)
        
        if not all_capacities:
            continue
        
        from collections import Counter
        capacity_counts = Counter([round(c, 1) for c in all_capacities])
        most_common_cap = capacity_counts.most_common(1)[0][0]
        
        if abs(current_q - most_common_cap) > 0.5:
            new_liters = round(most_common_cap * 0.946, 1)
            multi_engine_data[eng_str]["oil_capacity"]["with_filter"] = {
                "quarts": round(most_common_cap, 2),
                "liters": new_liters
            }
            print(f"      CORRECTED: {eng_str} capacity {current_q}qt -> {round(most_common_cap, 2)}qt")
    
    return multi_engine_data


def extract_all_capacities_for_engine(doc, engine_str):
    """Find oil-only capacity values near a specific engine size."""
    capacities = []
    
    match = re.search(r'(\d+\.?\d*)', engine_str)
    if not match:
        return capacities
    
    eng_num = match.group(1)
    
    for page in doc:
        text = page.get_text()
        text_lower = text.lower()
        
        oil_keywords = r"(?:engine\s+oil|oil\s+with\s+filter|oil\s+capacity|oil\s+change|oil\s+drain|oil\s+fill)"
        
        for match in re.finditer(oil_keywords, text_lower, re.IGNORECASE):
            start = max(0, match.start() - 50)
            end = min(len(text), match.end() + 200)
            window = text[start:end]
            
            if eng_num.replace(".", r"\.") in window or eng_num in window.lower():
                for cap_match in re.finditer(CAPACITY_PATTERN, window, re.IGNORECASE):
                    try:
                        q, l = to_quarts_liters(cap_match.group(1), cap_match.group(2))
                        if ENGINE_OIL_CAPACITY_MIN_QT <= q <= ENGINE_OIL_CAPACITY_MAX_QT:
                            capacities.append(q)
                    except (ValueError, TypeError):
                        continue
    
    return capacities


def detect_all_engines_in_pdf(doc):
    """Fallback scanner for all engine sizes mentioned anywhere in the PDF."""
    all_engines = []
    full_text = ""
    
    for page in doc:
        full_text += page.get_text().lower() + "\n"
    
    engine_matches = []
    for m in re.finditer(ENGINE_PATTERN, full_text):
        displacement = float(m.group(1))
        if not is_plausible_engine_displacement(displacement):
            continue
        eng_size = f"{displacement:.1f}L"
        start = max(0, m.start() - 100)
        end = min(len(full_text), m.end() + 200)
        context = full_text[start:end]
        
        engine_matches.append((eng_size, context, m.start()))
    
    seen = set()
    for eng_size, context, pos in engine_matches:
        variant = extract_engine_variant_from_context(context, eng_size)
        
        full_eng = eng_size + variant
        if full_eng not in seen:
            all_engines.append(full_eng)
            seen.add(full_eng)
    
    return all_engines


def fix_unknown_engine(doc, engine_data):
    """Replace unknown_engine with the best detected engine when the capacity clearly matches."""
    if "unknown_engine" not in engine_data:
        return engine_data
    
    unknown_cap = engine_data["unknown_engine"]["oil_capacity"]["with_filter"]
    if not unknown_cap:
        return engine_data
    
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
    
    if best_match:
        engine_data[best_match] = engine_data.pop("unknown_engine")
    
    return engine_data


def extract_engine_capacities(doc):
    """Extract engine-oil capacities from explicit oil sections and ignore other fluid tables."""
    engine_caps = {}
    
    full_text = ""
    for page in doc:
        full_text += page.get_text() + "\n"
    
    lines = full_text.split('\n')
    
    for line_idx, line in enumerate(lines):
        line_lower = line.lower()
        
        is_oil_header = any(kw in line_lower for kw in [
            'engine oil', 'motor oil', 'oil with filter', 'oil capacity', 'engine oil capacity'
        ])
        
        if not is_oil_header:
            continue
        
        # Oil tables are local: engine rows normally follow the oil header.
        oil_section_lines = lines[line_idx:min(line_idx + 15, len(lines))]
        
        for section_line_idx, section_line in enumerate(oil_section_lines):
            section_line_lower = section_line.lower()
            
            if len(section_line.strip()) < 3:
                continue

            if (
                section_line_idx > 0
                and re.search(r"\boil\s+filter\b", section_line_lower)
                and "engine oil" not in section_line_lower
                and "with filter" not in section_line_lower
            ):
                break
            
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
                    if (
                        re.search(r"\boil\s+filter\b", next_line_lower)
                        and "engine oil" not in next_line_lower
                        and "with filter" not in next_line_lower
                    ):
                        break
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
            
            filtered_engine_matches = []
            for em in engine_matches:
                context_before = section_line_lower[max(0, em.start() - 20):em.start()]
                if (
                    not is_plausible_engine_displacement(float(em.group(1)))
                    or
                    re.search(r"\d+\.?\d*\s*" + CAPACITY_UNIT_PATTERN + r"\s*\(?\s*$", context_before, re.I)
                    or is_capacity_conversion_engine_match(section_line_lower, em)
                    or is_parenthesized_capacity_conversion(section_line_lower, em)
                    or overlaps_real_capacity_match(section_line_lower, em, capacity_matches)
                ):
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
                    if ENGINE_OIL_CAPACITY_MIN_QT <= q <= ENGINE_OIL_CAPACITY_MAX_QT and "unknown_engine" not in engine_caps:
                        engine_caps["unknown_engine"] = {
                            "with_filter": {"quarts": q, "liters": l},
                            "without_filter": None
                        }
                except (ValueError, TypeError, AttributeError):
                    pass

                continue
            
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
                
                selected_cap = None
                for cm in capacity_matches:
                    if cm.start() >= eng_match.end():
                        unit = cm.group(2).lower()
                        if 'qt' in unit or 'quart' in unit:
                            selected_cap = cm
                            break
                
                if not selected_cap:
                    for cm in capacity_matches:
                        if cm.start() >= eng_match.end():
                            selected_cap = cm
                            break
                
                if not selected_cap:
                    for next_line in oil_section_lines[section_line_idx + 1:section_line_idx + 4]:
                        next_line_lower = next_line.lower()
                        if (
                            re.search(r"\boil\s+filter\b", next_line_lower)
                            and "engine oil" not in next_line_lower
                            and "with filter" not in next_line_lower
                        ):
                            break
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
                
                if ENGINE_OIL_CAPACITY_MIN_QT <= q <= ENGINE_OIL_CAPACITY_MAX_QT:
                    engine_caps[eng_size] = {
                        "with_filter": {"quarts": q, "liters": l},
                        "without_filter": None
                    }
            except (ValueError, TypeError, AttributeError):
                continue
    
    return engine_caps


def extract_explicit_shared_engine_oil_capacity(doc):
    """Find shared rows like "Engine Oil with Filter 8.0 L 8.5 qt" before broad fallbacks."""
    full_text = ""
    for page in doc:
        full_text += page.get_text() + "\n"

    lines = full_text.split("\n")
    for line_idx, line in enumerate(lines):
        line_lower = line.lower()
        if "engine oil" not in line_lower:
            continue
        if "filter" not in line_lower and "capacity" not in line_lower:
            continue
        if any(term in line_lower for term in ["oil filter", "part number", "acdelco"]):
            continue

        window_lines = [line]
        for next_line in lines[line_idx + 1:line_idx + 4]:
            next_lower = next_line.lower()
            if re.search(r"\boil\s+filter\b", next_lower) and "engine oil" not in next_lower:
                break
            if any(term in next_lower for term in [
                "fuel tank", "cooling system", "transfer case", "wheel nut",
                "transmission", "differential", "refrigerant"
            ]):
                break
            window_lines.append(next_line)

        window = " ".join(window_lines).lower()
        capacity_matches = [
            m for m in re.finditer(CAPACITY_PATTERN, window, re.I)
            if is_real_capacity_match(window, m)
        ]
        if not capacity_matches:
            continue

        selected = None
        for match in capacity_matches:
            unit = match.group(2).lower()
            if "qt" in unit or "quart" in unit:
                selected = match
                break
        if not selected:
            selected = capacity_matches[0]

        try:
            q, l = to_quarts_liters(selected.group(1), selected.group(2))
        except (ValueError, TypeError, AttributeError):
            continue

        if ENGINE_OIL_CAPACITY_MIN_QT <= q <= ENGINE_OIL_CAPACITY_MAX_QT:
            return {"with_filter": {"quarts": q, "liters": l}, "without_filter": None}

    return None


def extract_fallback_capacity(doc):
    """Find a generic oil capacity when no engine-specific capacity can be extracted."""
    explicit_shared = extract_explicit_shared_engine_oil_capacity(doc)
    if explicit_shared:
        return explicit_shared

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
            if ENGINE_OIL_CAPACITY_MIN_QT <= q <= ENGINE_OIL_CAPACITY_MAX_QT:
                return {"with_filter": {"quarts": q, "liters": l}, "without_filter": None}
        
        for m in re.finditer(CAPACITY_PATTERN, text):
            if not is_real_capacity_match(text, m):
                continue
            q, l = to_quarts_liters(m.group(1), m.group(2))
            if ENGINE_OIL_CAPACITY_MIN_QT <= q <= ENGINE_OIL_CAPACITY_MAX_QT:
                return {"with_filter": {"quarts": q, "liters": l}, "without_filter": None}
    
    return None

def extract_oils(text):
    """
    Extract oil viscosities, recommendation strength, temperature conditions, and inline
    engine links.
    """
    oil_scores = {}
    oil_temps = {}
    engine_oil_map_inline = {}
    do_not_use_oils = set()
    
    # Pass 0: remove oils that are explicitly unsuitable, while keeping
    # conditional alternatives such as "0W-30 may be used".
    do_not_use_pattern = r"do\s+not\s+use|should\s+not\s+be\s+used"
    for match in re.finditer(do_not_use_pattern, text.lower()):
        start_pos = match.start()
        
        end_pos = start_pos + 500
        for i in range(match.end(), min(len(text), match.end() + 500)):
            if (text[i] == '.' or text[i] == '\n') and i + 1 < len(text):
                j = i + 1
                while j < len(text) and text[j] in ' \n\t':
                    j += 1
                if j < len(text) and (text[j].isupper() or text[j:j+3].upper() == 'SAE'):
                    end_pos = i
                    break
        
        window = text[start_pos:end_pos]
        
        if "may be used" in window.lower() or "can be used" in window.lower() or "alternative" in window.lower():
            continue
            
        for oil_match in re.finditer(OIL_PATTERN, window.lower(), re.I):
            if oil_match:
                base, grade = oil_match.groups()[-2:]
                oil = normalize_oil(f"{base}W-{grade}")
                do_not_use_oils.add(oil)
    
    # Pass 1: engine rows sometimes carry their own "(SAE 5W-30)" oil note.
    lines = text.split('\n')
    for line in lines:
        line_lower = line.lower()
        if re.search(ENGINE_PATTERN, line_lower) and '(sae' in line_lower:
            eng_match = re.search(ENGINE_PATTERN, line_lower)
            if eng_match and is_plausible_engine_displacement(float(eng_match.group(1))):
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
                        temps = extract_temperature(line)
                        final_temps = get_temperature_with_fallback(temps, oil)
                        oil_temps[oil].update(final_temps)
    
    # Pass 2: "best viscosity grade" is the strongest primary-oil signal.
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
                
                sentence_start = match.start()
                sentence_end = min(len(text), match.end() + 300)
                
                # Limit context to one statement so nearby alternatives are not merged.
                for i in range(match.end(), min(len(text), match.end() + 300)):
                    if (text[i] == '.' or text[i] in '\n') and i + 1 < len(text):
                        j = i + 1
                        while j < len(text) and text[j] in ' \n\t':
                            j += 1
                        if j < len(text) and (text[j].isupper() or text[j:j+3].upper() == 'SAE'):
                            sentence_end = i
                            break
                
                context = text[sentence_start:sentence_end].lower()
                
                all_engines_in_context = []
                for eng_m in re.finditer(ENGINE_PATTERN, context):
                    if not is_plausible_engine_displacement(float(eng_m.group(1))):
                        continue
                    eng_size = f"{float(eng_m.group(1)):.1f}L"
                    all_engines_in_context.append(eng_size)
                    if eng_size not in engine_oil_map_inline:
                        engine_oil_map_inline[eng_size] = []
                    if oil not in engine_oil_map_inline[eng_size]:
                        engine_oil_map_inline[eng_size].append(oil)
                
                if not oil_temps[oil]:
                    oil_temps[oil].add("all temperatures")

    # Pass 2.25: catch recommended/preferred wording not covered by "best".
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
    
    # Pass 2.5: capture conditional oils, often in cold-temperature sections.
    conditional_oil_pattern = r'(?:an?\s+)?sae\s+(0|5|10|15|20|25)w[-\u2013\u2014]?\s*(\d+)(?:\s+oil)?\s+(?:may\s+be\s+used|can\s+be\s+used|is\s+acceptable)'
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
            
            oil_scores[oil] += 2
            
            temps = extract_temperature(context)
            
            final_temps = get_temperature_with_fallback(temps, oil)
            oil_temps[oil].update(final_temps)
            
            # If the conditional statement names an engine, link the oil to it.
            rel_match_start = match.start() - context_start
            statement_context = context_lower[rel_match_start:min(len(context_lower), rel_match_start + 200)]

            for eng_m in re.finditer(ENGINE_PATTERN, statement_context):
                if not is_plausible_engine_displacement(float(eng_m.group(1))):
                    continue
                eng_size = f"{float(eng_m.group(1)):.1f}L"
                if eng_size not in engine_oil_map_inline:
                    engine_oil_map_inline[eng_size] = []
                if oil not in engine_oil_map_inline[eng_size]:
                    engine_oil_map_inline[eng_size].append(oil)
    
    # Pass 3: sentence-level sweep for remaining engine-oil recommendations.
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
            temps.add("above 20F (-7C)")
        elif "year-round" in lower or "all temperatures" in lower:
            temps.add("all temperatures")
        else:
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
            
            final_temps = get_temperature_with_fallback(temps, oil)
            oil_temps[oil].update(final_temps)
    
    # Pass 4: final proximity scan for oil mentions in valid engine-oil context.
    all_oils = re.findall(OIL_PATTERN, text)
    for base, grade in all_oils:
        oil = normalize_oil(f"{base}W-{grade}")
        if oil in do_not_use_oils:
            continue
        if oil not in oil_scores:
            pattern = f"{base}\\s*W[-\\u2013\\u2014]?\\s*{grade}"
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
    
    # Apply document-wide temperatures only to oils that did not get local context.
    doc_temps = extract_temperature(text)
    has_doc_specific_temps = any("F" in t or "C" in t or "range:" in t for t in doc_temps if t != "all temperatures")
    
    if has_doc_specific_temps:
        for oil in oil_scores:
            if "unknown" in oil_temps.get(oil, set()):
                oil_temps[oil] = doc_temps
    elif "never goes below" in text.lower():
        for oil in oil_scores:
            if "unknown" in oil_temps.get(oil, set()):
                oil_temps[oil] = {"above 20F (-7C)"}
    elif "year-round" in text.lower():
        for oil in oil_scores:
            if "unknown" in oil_temps.get(oil, set()):
                oil_temps[oil] = {"all temperatures"}
    
    return oil_scores, oil_temps, engine_oil_map_inline

def select_best_engine(engine_caps, all_engines):
    """Pick a representative engine/capacity pair for fallback logic."""
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
        
        realistic = []
        for e in sorted_engines:
            disp = get_engine_displacement(e)
            if is_plausible_engine_displacement(disp):
                realistic.append(e)
        
        best = realistic[0] if realistic else sorted_engines[0]
        return best, None

    return None, None

def extract_all():
    """Run the full Drive-to-JSON extraction pipeline."""
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
            extraction_type, avg_chars = analyze_pdf_type(doc)
            
            if extraction_type == "AUTO":
                print(f"  Using text extraction ({avg_chars} chars/page - text-heavy PDF)")
            else:
                print(f"  Using OCR extraction ({avg_chars} chars/page - scanned/manual document)")
            
            if year is None:
                y2, m2, mo2 = detect_vehicle_from_pdf(doc)
                year  = year  or y2
                make  = make  or m2
                model = model or mo2

            pages_with_images = []
            text_parts = []
            for page_num, p in enumerate(doc):
                page_text = p.get_text("text")
                text_parts.append(page_text)
                
                if extraction_type == "MANUAL":
                    images = p.get_images()
                    if images:
                        pages_with_images.append(page_num)
            
            # Raw text preserves table rows; cleaned text is better for prose scans.
            raw_text = "\n".join(text_parts)
            text = clean_text(raw_text)
            
            if extraction_type == "MANUAL" and pages_with_images:
                print(f"      Running OCR on {len(pages_with_images)} image page(s)...")
                ocr_text = extract_text_from_images(doc, pages_with_images)
                if ocr_text.strip():
                    raw_text = raw_text + "\n" + ocr_text
                    text = clean_text(raw_text)
                    print(f"      OCR extracted {len(ocr_text)} characters")
            
            # Engine evidence priority: structured tables first, prose fallback second.
            table_engines = extract_engines_from_spec_table(raw_text)
            if table_engines:
                print(f"      TABLE EXTRACTION found engines: {table_engines}")
            
            if table_engines:
                all_engines = table_engines
            else:
                all_engines = extract_engines(text)
                if all_engines:
                    print(f"      General extraction found engines: {all_engines}")
            
            all_engines = filter_engine_outliers(all_engines)
            all_engines = consolidate_engine_variants(all_engines)
            
            engine_caps = extract_engine_capacities(doc)
            shared_fallback_cap = extract_fallback_capacity(doc)
            engine_caps = prefer_shared_capacity_if_current_caps_are_noise(engine_caps, shared_fallback_cap)
            if all_engines and engine_caps:
                engine_caps = filter_engine_caps_to_detected_engines(engine_caps, all_engines)
            
            # When oil capacities are available, they validate which detected
            # engines belong in the final output.
            if all_engines and engine_caps:
                valid_engine_bases = set()
                for cap_eng in engine_caps.keys():
                    valid_engine_bases.add(cap_eng.split()[0])
                
                filtered_engines = []
                for eng in all_engines:
                    eng_base = eng.split()[0]
                    if eng_base in valid_engine_bases:
                        filtered_engines.append(eng)
                
                if filtered_engines:
                    all_engines = filtered_engines

            # Engine-specific oil rows can add missing engines. Shared-capacity
            # rows stay as unknown_engine until expanded to already detected engines.
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
                specific_temps = {t for t in temps if "F" in t or "C" in t or "range:" in t or "above" in t or "below" in t}
                if specific_temps and "all temperatures" in temps:
                    oil_temps[oil] = specific_temps
            
            engine_types = []
            if all_engines:
                for eng in all_engines:
                    parts = eng.split()
                    for part in parts[1:]:
                        part_upper = part.upper()
                        if part_upper not in engine_types and part_upper not in ["L", "LITER", "LITRE"]:
                            engine_types.append(part_upper)
            
            engine_types_from_text = extract_engine_types(text, all_engines)
            for et in engine_types_from_text:
                if et not in engine_types:
                    engine_types.append(et)
            
            engine_oil_map = map_oils_to_engines(text)
            
            # Inline engine-oil links are more specific than proximity links.
            for eng_size, oils_list in engine_oil_map_inline.items():
                if eng_size not in engine_oil_map:
                    engine_oil_map[eng_size] = oils_list[:]
                else:
                    for oil in oils_list:
                        if oil not in engine_oil_map[eng_size]:
                            engine_oil_map[eng_size].insert(0, oil)
            
            # If the strict parser missed a shared oil capacity, try the
            # oil-specific fallback before broad numeric pairing.
            if not engine_caps:
                if shared_fallback_cap:
                    engine_caps["unknown_engine"] = shared_fallback_cap

            # Only use generic capacity pairing as the final resort. It scans
            # more broadly, so oil-section extraction above is safer.
            if not engine_caps and all_engines:
                caps = []

                for m in re.finditer(CAPACITY_PATTERN, text):
                    if not is_real_capacity_match(text, m):
                        continue
                    q, l = to_quarts_liters(m.group(1), m.group(2))
                    if ENGINE_OIL_CAPACITY_MIN_QT <= q <= ENGINE_OIL_CAPACITY_MAX_QT:
                        caps.append({"quarts": q, "liters": l})

                paired_caps = pair_quarts_liters(caps)

                for i in range(min(len(all_engines), len(paired_caps))):
                    engine_caps[all_engines[i]] = {
                        "with_filter": paired_caps[i],
                        "without_filter": None
                    }

            engine_caps = expand_shared_capacity_to_detected_engines(engine_caps, all_engines)
            engine_caps = align_capacity_engine_keys_with_detected_variants(engine_caps, all_engines)
            engine_caps = filter_engine_caps_to_detected_engines(engine_caps, all_engines)
            
            
            multi_engine_data = build_multi_engine_data(engine_caps, oil_scores, oil_temps, engine_oil_map)
            
            # Keep the correction pass disabled; broader capacity rescans tended
            # to mix oil with coolant/fuel/transmission tables.
            # multi_engine_data = apply_correct_capacities(multi_engine_data, doc)
            
            if multi_engine_data:
                for eng_str, eng_info in multi_engine_data.items():
                    cap_info = eng_info.get("oil_capacity", {})
                    with_filter = cap_info.get("with_filter")
                    without_filter = cap_info.get("without_filter")
                    
                    if with_filter and "quarts" in with_filter:
                        with_filter["quarts"] = normalize_capacity_value(with_filter["quarts"])
                    if without_filter and "quarts" in without_filter:
                        without_filter["quarts"] = normalize_capacity_value(without_filter["quarts"])
                
                multi_engine_data = fix_unknown_engine(doc, multi_engine_data)

                # Final sanity filter: output engines must match extracted engine bases.
                if all_engines and multi_engine_data:
                    valid_engine_bases = {eng.split()[0] for eng in all_engines}
                    filtered_multi_engine_data = {
                        eng_key: eng_val
                        for eng_key, eng_val in multi_engine_data.items()
                        if eng_key.split()[0] in valid_engine_bases
                    }
                    if filtered_multi_engine_data:
                        multi_engine_data = filtered_multi_engine_data

                multi_engine_data = apply_shared_capacity_to_noisy_engine_data(
                    multi_engine_data,
                    shared_fallback_cap
                )
                
                for eng_str, eng_info in multi_engine_data.items():
                    cap_info = eng_info.get("oil_capacity", {})
                    with_filter = cap_info.get("with_filter")
                    if with_filter and "quarts" in with_filter:
                        cap_q = with_filter["quarts"]
                        if not (ENGINE_OIL_CAPACITY_MIN_QT <= cap_q <= ENGINE_OIL_CAPACITY_MAX_QT):
                            print(f"      Warning: {eng_str} capacity {cap_q}qt seems outside normal automotive oil range")

                multi_engine_data = add_missing_engine_type_to_keys(multi_engine_data, engine_types)
            
            if not multi_engine_data:
                fallback_cap = extract_fallback_capacity(doc)

                oil_list = []
                if oil_scores:
                    primary = max(oil_scores, key=oil_scores.get)
                    max_score = oil_scores[primary]
                    
                    for oil, score in oil_scores.items():
                        temps = oil_temps.get(oil, [])
                        has_actual_temps = any(
                            ("F" in t or "C" in t or 
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

                multi_engine_data = add_missing_engine_type_to_keys(multi_engine_data, engine_types)
            
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
