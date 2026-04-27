import fitz
import re
import io
import json
import time
from pathlib import Path
from PIL import Image
import pytesseract
from PIL import ImageFilter, ImageOps
from googleapiclient.discovery import build
from google.oauth2 import service_account
from googleapiclient.http import MediaIoBaseDownload


FOLDER_ID = "1J22Hv9BJD5AoB-jCepMMQgEriM-eIVnq"
SCOPES = ['https://www.googleapis.com/auth/drive.readonly']
CREDS_PATH = "credentials.json"
OUTPUT_FILE = "structured_results.json"
REFERENCE_FILE = "vehicle_reference.json"
LOCAL_MANUALS_FOLDER = "Manuals"


# Core text patterns. These stay generic: they describe syntax commonly found in
# manuals, not any specific vehicle. Unicode escapes allow PDF dash/minus/degree
# variants without putting corrupted characters in the source file.
OIL_PATTERN = r"\b(0|5|10|15|20|25)\s*W[-\u2013\u2014]?\s*(16|20|30|40|50|60)\b"
CAPACITY_PATTERN = r"(\d+\.?\d*)\s*(?:us\s+|imp\s+|u\.s\.\s+)?(quarts?|qts?|qt\.?|gallons?|gal\.?|liters?|litres?|l\b)"
CAPACITY_UNIT_PATTERN = r"(?:quarts?|qts?|qt\.?|liters?|litres?|l\b|gal|gallons?|ml|cc)"
ENGINE_PATTERN = r"\b(\d{1,2}\.\d)\s*(?:(?:[-]?\s*(?:l|liter|litre)(?=\b|(?:v|i|l)\s*-?\d))|(?:\s+(?:gdi|dohc|sohc|turbo|ecoboost|naturally|cylinder)))"
ENGINE_CODE_PATTERN = r"\b(?:vortec[\W_]{0,4})?([1-9]\d{3})\s*(?:(supercharged|turbocharged|turbo)\s+)?(?:(?:series|serie|gen|generation)\s*(?:i{1,3}|iv|v|1|2|3|4|5)\s+)?(?:sfi\s+|mpi\s+|efi\s+)?((?:v|i|l)\s*-?\s*(?:4|5|6|8|10|12)|inline\s*-?\s*(?:4|5|6|8))\b"
ENGINE_TYPE_PATTERN = r"\b(?:v\s*-?\s*(?:3|4|5|6|8|10|12|16|20|24)|i\s*-?\s*[3-8]|l\s*-?\s*[3-8]|w\s*-?\s*(?:8|12|16)|h\s*-?\s*4|f\s*-?\s*8|inline\s*-?\s*[3-8]|flat\s*-?\s*(?:3|4|6|8|12)|boxer|rotary|wankel|turbo(?:charged)?|supercharged|naturally\s*-?\s*aspirated|ecoboost|hybrid|electric|diesel|petrol|gdi|sohc|dohc)\b"
CYLINDER_WORDS = {
    "three": "3",
    "four": "4",
    "five": "5",
    "six": "6",
    "seven": "7",
    "eight": "8",
}
CYLINDER_PATTERN = r"\b(?:(3|4|5|6|7|8)|three|four|five|six|seven|eight)\s*[-]?\s*cylinder\b"
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

NON_VEHICLE_ENGINE_CONTEXT = [
    "warranty", "coverage", "covered", "guide", "eligible", "emissions",
    "defect", "performance warranties", "service data", "recording",
    "diesel engine coverage", "new vehicle limited warranty"
]

VEHICLE_NAVIGATION_STOP_WORDS = {
    "section", "sections", "how", "features", "controls", "control",
    "audio", "comfort", "road", "problems", "appearance", "care",
    "maintenance", "schedule", "customer", "assistance", "index",
    "summary", "contents", "restraint", "systems", "driving", "seats"
}

OIL_PAGE_KEYWORDS = [
    "oil capacity", "engine oil", "crankcase", "oil with filter",
    "oil change capacity", "including filter", "engine oil capacity",
    "engine oil recommendation", "viscosity", "api service",
    "lubricant", "specifications", "technical information"
]

CAPACITY_REFILL_CONTEXT = [
    "refill",
    "add 1",
    "add one",
    "add engine oil",
    "top off",
    "top-up",
    "top up",
    "oil level",
    "at next refueling",
    "within the next",
    "minimum oil warning",
    "check engine oil level",
    "engine oil level refill",
    "engine oil level reduce",
    "engine oil level stop",
    "up to 1 us quart",
    "up to 1 quart",
    "up to 1 l",
    "up to 1 liter",
    "up to 1 litre",
    "consumption may be",
]

NON_CAPACITY_OIL_QUANTITY_CONTEXT = [
    "do not use more than",
    "not use more than",
    "between scheduled service intervals",
    "between service intervals",
    "alternative engine oil",
    "raise the indicated level",
    "minimum to maximum",
    "required to raise",
    "oil level",
    "dipstick",
    "adding engine oil",
    "add oil",
    "top off",
    "top-up",
    "top up",
    "oil consumption",
]

ENGINE_OIL_STOP_TERMS = [
    "cooling system", "coolant", "fuel tank", "transmission",
    "transaxle", "differential", "power steering", "brake fluid",
    "washer", "refrigerant", "wheel nut", "transfer case"
]

WITH_FILTER_LABELS = [
    "including the oil filter", "includes filter change", "includes filter",
    "including filter", "with oil filter", "with filter", "variant including"
]

WITHOUT_FILTER_LABELS = [
    "excluding the oil filter", "without filter"
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
    if normalized == "ECOBOOST":
        return "ECOBOOST"
    if normalized == "NATURALLYASPIRATED":
        return "NATURALLY-ASPIRATED"

    return normalized


def format_engine_variant_token(token):
    """Format normalized variant tokens for engine labels."""
    for family in load_engine_family_reference():
        if normalize_engine_type_token(family) == token:
            return family
    if token == "ECOBOOST":
        return "EcoBoost"
    if token == "NATURALLY-ASPIRATED":
        return "Naturally Aspirated"
    if token in {"TURBO", "SUPERCHARGED"}:
        return token.title()
    return token


def normalize_cylinder_match(match):
    """Convert numeric or word cylinder text to an inline layout token."""
    digit = match.group(1)
    if not digit:
        digit = CYLINDER_WORDS.get(match.group(0).split("-")[0].split()[0].lower())
    return f"I{digit}" if digit else ""


def find_engine_family_tokens(context):
    """Find reference engine-family labels such as EcoBoost, Hemi, or Skyactiv-G."""
    if not context:
        return []

    found = []
    matched_spans = []
    families = sorted(load_engine_family_reference(), key=len, reverse=True)
    for family in families:
        pattern = r"\b" + re.escape(family).replace(r"\ ", r"\s+") + r"\b"
        for match in re.finditer(pattern, context, re.I):
            overlaps_existing = any(
                match.start() < span_end and span_start < match.end()
                for span_start, span_end in matched_spans
            )
            if overlaps_existing:
                continue

            token = normalize_engine_type_token(family)
            if token and token not in found:
                found.append(token)
            matched_spans.append((match.start(), match.end()))
            break

    return found


def extract_engine_variant_from_context(context, base_engine=""):
    """Find an engine layout/variant token near a displacement."""
    if not context:
        return ""

    base_upper = base_engine.upper()
    tokens = []

    for match in re.finditer(r"\b\d{1,2}\.\d\s*l\s*((?:v|i|l)\s*-?\s*(?:3|4|5|6|8|10|12))\b", context, re.I):
        token = normalize_engine_type_token(match.group(1))
        if token and token not in tokens and token not in base_upper:
            tokens.append(token)

    for match in re.finditer(ENGINE_TYPE_PATTERN, context, re.I):
        token = normalize_engine_type_token(match.group())
        if not token or token in base_upper:
            continue
        if token in {"GDI", "SOHC", "DOHC", "DIESEL", "PETROL", "HYBRID", "ELECTRIC"}:
            continue
        if token not in tokens:
            tokens.append(token)

    for token in find_engine_family_tokens(context):
        if token and token not in tokens and token not in base_upper:
            tokens.append(token)

    for match in re.finditer(CYLINDER_PATTERN, context, re.I):
        token = normalize_cylinder_match(match)
        if token and token not in tokens and token not in base_upper:
            tokens.append(token)

    if not tokens:
        return ""

    layout_tokens = [t for t in tokens if re.fullmatch(r"(?:I|V|W|H|F)\d{1,2}", t)]
    variant_tokens = [t for t in tokens if t not in layout_tokens]
    ordered_tokens = variant_tokens + layout_tokens
    return " " + " ".join(format_engine_variant_token(t) for t in ordered_tokens)


def normalize_engine_code_label(match, context=""):
    """Convert engine-family codes like 3800 V6 or Vortec 5300 V8 to 3.8L V6."""
    if not match:
        return ""

    try:
        displacement = int(match.group(1)) / 1000.0
    except (TypeError, ValueError):
        return ""

    if not is_plausible_engine_displacement(displacement):
        return ""

    context_lower = str(context).lower()
    if any(term in context_lower for term in NON_VEHICLE_ENGINE_CONTEXT):
        return ""
    if is_capacity_or_fluid_row(context_lower) and "engine oil" not in context_lower:
        return ""

    layout = normalize_engine_type_token(match.group(3))
    if not is_layout_engine_type(layout):
        return ""

    forced_variant = normalize_engine_type_token(match.group(2) or "")
    base_engine = f"{displacement:.1f}L"
    variant_tokens = []
    if forced_variant in {"SUPERCHARGED", "TURBO"}:
        variant_tokens.append(format_engine_variant_token(forced_variant))

    return " ".join([base_engine] + variant_tokens + [layout])


def extract_engine_code_labels(text):
    """Find code-style engine labels used by older manuals, e.g. 3100 V6."""
    labels = []
    if not text:
        return labels

    for match in re.finditer(ENGINE_CODE_PATTERN, text, re.I):
        start = max(0, match.start() - 80)
        end = min(len(text), match.end() + 120)
        context = text[start:end]
        label = normalize_engine_code_label(match, context)
        if label and label not in labels:
            labels.append(label)

    return labels


def engine_code_aliases(engine_str):
    """Return compact aliases for a decimal engine label, such as 3.8L -> 3800."""
    displacement = get_engine_displacement(engine_str)
    if not is_plausible_engine_displacement(displacement):
        return []

    aliases = {f"{displacement:.1f}", str(int(round(displacement * 1000)))}
    return sorted(aliases, key=len, reverse=True)


def has_engine_signal(text):
    """Return True when a line has enough context to be treated as engine-related text."""
    if not text:
        return False

    text_lower = text.lower()

    if any(term in text_lower for term in NON_VEHICLE_ENGINE_CONTEXT):
        return False

    # Primary signals: actual engine sizes and engine type tokens.
    if re.search(ENGINE_PATTERN, text_lower):
        return True
    if re.search(r"\b(?:[a-z]+\s+)?g?\d{1,2}\.\d\s+(?:t-?gdi|tgdi|gdi|mpi|crdi|diesel|turbo)\b", text_lower, re.I):
        return True
    if re.search(ENGINE_TYPE_PATTERN, text_lower):
        return True
    if find_engine_family_tokens(text_lower):
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

    paired_with_us_capacity = (
        re.search(r"\d+\.?\d*\s*(?:quarts?|qts?|qt\.?|gallons?|gal\.?)\s*\(?\s*$", before, re.I)
        or re.search(r"^\s*\)?\s*\d+\.?\d*\s*(?:quarts?|qts?|qt\.?|gallons?|gal\.?)", after, re.I)
    )
    if paired_with_us_capacity:
        return True

    if re.search(r"^\s*(?:engine|engines|v\d|i\d|turbo|flex fuel)", after, re.I):
        return False

    if unit.startswith('l'):
        return False

    return any(term in text.lower() for term in [
        "capacity", "capacities", "with filter", "fluid", "cooling system",
        "fuel tank", "quarts", "qts", "qt"
    ])


def match_sentence_context(text, match, radius=180):
    """Return the sentence-like context around a numeric quantity match."""
    if not text or not match:
        return ""

    raw = str(text)
    start_candidates = [
        raw.rfind(delimiter, 0, match.start())
        for delimiter in (".", ";", "!", "?")
    ]
    start = max(start_candidates)
    if start == -1:
        start = max(0, match.start() - radius)
    else:
        start += 1

    end_candidates = [
        raw.find(delimiter, match.end())
        for delimiter in (".", ";", "!", "?")
    ]
    end_candidates = [idx for idx in end_candidates if idx != -1]
    end = min(end_candidates) + 1 if end_candidates else min(len(raw), match.end() + radius)

    return raw[start:end].lower()


def is_non_capacity_oil_quantity_match(text, match):
    """Reject oil-service quantities that are not total engine-oil capacities."""
    local = match_sentence_context(text, match)
    if not local:
        return False

    if any(term in local for term in NON_CAPACITY_OIL_QUANTITY_CONTEXT):
        return True
    if re.search(r"quantity\s+of\s+engine\s+oil.{0,80}raise.{0,80}dipstick", local, re.I):
        return True
    if re.search(r"do\s+not\s+use\s+more\s+than.{0,80}engine\s+oil", local, re.I):
        return True
    if re.search(r"between\s+(?:scheduled\s+)?service\s+intervals", local, re.I):
        return True

    return False


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
        "what kind of oil to use",
        "proper viscosity oil for your vehicle",
        "recommended sae viscosity grade engine oils",
        "next oil change",
        "synthetic oil",
        "very cold",
        "cold starting",
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


def score_oil_evidence(text, oil=None):
    """Score an oil recommendation candidate from its local statement/context."""
    text_lower = str(text).lower()
    score = 0

    if has_non_engine_oil_context(text_lower):
        return -10
    if has_engine_oil_context(text_lower):
        score += 6

    if "best viscosity" in text_lower:
        score += 8
    if "recommended" in text_lower or "preferred" in text_lower:
        score += 5
    if (
        "may be used" in text_lower
        or "can be used" in text_lower
        or "can use" in text_lower
        or "you can use" in text_lower
        or "consider using" in text_lower
        or "acceptable" in text_lower
    ):
        score += 2
    if "year-round" in text_lower or "all temperatures" in text_lower:
        score += 2
    if "do not use" in text_lower or "should not be used" in text_lower:
        score -= 12
    if oil and oil.lower() in text_lower and "sae" in text_lower:
        score += 1

    return score



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


def extract_targeted_spec_ocr_text(doc):
    """OCR image-backed spec/capacity pages even when the whole PDF is text-heavy."""
    if not doc:
        return ""

    target_pages = set()
    for page_num, page in enumerate(doc):
        page_text = page.get_text("text")
        page_lower = page_text.lower()
        page_lines = [re.sub(r"\s+", " ", line).strip().lower() for line in page_text.splitlines()]
        has_spec_heading = any(
            line in {
                "capacities and specifications",
                "engine specifications",
                "engine specification",
                "refill capacities",
                "lubricant specifications",
                "engine data",
            }
            for line in page_lines
        )
        has_capacity_table_text = (
            "capacities and specifications" in page_lower
            and (
                "application" in page_lower
                or "refill capacities" in page_lower
                or "capacity" in page_lower
                or "engine data" in page_lower
            )
            and ("capacities" in page_lower or "capacity" in page_lower)
        )
        has_dotted_index_entry = bool(re.search(
            r"(?:capacities and specifications|engine specifications)\s*\.{3,}",
            page_lower,
            re.I,
        ))
        first_page_text = " ".join(page_lines[:12])
        has_contents_navigation = (
            (
                "contents" in first_page_text
                or sum(
                    1
                    for term in [
                        "introduction",
                        "instrumentation",
                        "controls",
                        "seating",
                        "roadside",
                        "customer assistance",
                        "index",
                    ]
                    if term in page_lower
                ) >= 4
            )
            and "capacities and specifications" in page_lower
            and not page_lower.lstrip().startswith("capacities and specifications")
        )
        has_relevant_text_table = bool(page.get_images()) or any(
            term in page_lower
            for term in [
                "refill capacities",
                "lubricant specifications",
                "engine data",
                "motorcraft part numbers",
                "application",
                "capacity",
            ]
        )

        if (
            (has_spec_heading or has_capacity_table_text)
            and has_relevant_text_table
            and not has_dotted_index_entry
            and not has_contents_navigation
        ):
            target_pages.add(page_num)
            if page.get_images() and page_num + 1 < len(doc):
                target_pages.add(page_num + 1)

    ocr_parts = []
    for page_num in sorted(target_pages):
        try:
            page = doc[page_num]
            pix = page.get_pixmap(matrix=fitz.Matrix(8, 8), colorspace=fitz.csRGB)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            gray = ImageOps.grayscale(img)
            enhanced = gray.filter(ImageFilter.SHARPEN)
            page_ocr_parts = []
            for ocr_img, config in (
                (gray, "--psm 6"),
                (gray, "--psm 12"),
                (enhanced, "--psm 6"),
            ):
                text = pytesseract.image_to_string(ocr_img, config=config)
                if text.strip():
                    page_ocr_parts.append(text)
            if page_ocr_parts:
                ocr_parts.append("\n".join(page_ocr_parts))
        except Exception as e:
            print(f"      Warning: targeted spec OCR failed on page {page_num}: {str(e)}")
            continue

    return "\n".join(ocr_parts)


def normalize_capacity_ocr_text(text):
    """Clean OCR noise from capacity-table numeric cells."""
    normalized = str(text).lower()
    normalized = normalized.replace(",", ".")
    normalized = re.sub(r"(?<=\d)\s*\.\s*(?=\d)", ".", normalized)

    # OCR often drops the decimal point in table cells such as "4.5 quarts",
    # producing "45quans" or "4Squars". Keep this correction scoped to
    # capacity-row OCR, where whole-number 45 quart oil capacities are invalid.
    normalized = re.sub(
        r"\b([1-9])\s*[s5]\s*q[a-z]{2,8}\b",
        r"\1.5 quarts",
        normalized,
        flags=re.I,
    )
    normalized = re.sub(
        r"\b([1-9])\s*[s5]\s*(?:quart|quarts)\b",
        r"\1.5 quarts",
        normalized,
        flags=re.I,
    )
    normalized = re.sub(r"\bq[a-z]{2,8}\b", "quarts", normalized, flags=re.I)

    return normalized


def capacity_from_ocr_numeric_text(text):
    """Extract a plausible oil capacity from OCR text cropped to one table row."""
    normalized = normalize_capacity_ocr_text(text)
    matches = [
        match for match in re.finditer(CAPACITY_PATTERN, normalized, re.I)
        if is_real_capacity_match(normalized, match)
    ]

    selected = choose_preferred_capacity_match(matches)
    if selected:
        try:
            capacity = build_capacity_record_from_matches(selected, matches)
        except (ValueError, TypeError, AttributeError):
            capacity = None
        if capacity and ENGINE_OIL_CAPACITY_MIN_QT <= capacity.get("quarts", 0) <= ENGINE_OIL_CAPACITY_MAX_QT:
            return capacity

    for match in re.finditer(r"\b([1-9])([05])\s*(?:quarts?|q[a-z]{2,8})\b", normalized, re.I):
        try:
            quarts = float(f"{match.group(1)}.{match.group(2)}")
        except (TypeError, ValueError):
            continue
        if ENGINE_OIL_CAPACITY_MIN_QT <= quarts <= ENGINE_OIL_CAPACITY_MAX_QT:
            return {"quarts": quarts, "liters": round(quarts * 0.946352946, 1)}

    return None


def format_named_engine_variant(token):
    """Format named engine-family cells such as SPI or Zetec-E."""
    token = re.sub(r"[^a-z0-9-]+", "", str(token).strip(), flags=re.I)
    if not token:
        return ""
    if len(token) <= 3 or token.isupper():
        return token.upper()
    return "-".join(part[:1].upper() + part[1:].lower() for part in token.split("-") if part)


def extract_named_engine_variants_from_text(text):
    """Find labels like "2.0L SPI engine" or "2.0L Zetec-E engine" in spec tables."""
    if not text:
        return []

    generic_variant_words = {
        "engine", "engines", "gas", "gasoline", "petrol", "fuel", "oil",
        "coolant", "automatic", "manual", "standard", "optional", "base",
        "liter", "litre", "cylinder", "capacity", "system", "data",
    }
    variants = []
    seen = set()
    pattern = re.compile(
        r"\b(\d{1,2}\.\d)\s*l\s+([a-z][a-z0-9-]{1,20})\s+engine\b",
        re.I,
    )

    for match in pattern.finditer(str(text)):
        try:
            displacement = float(match.group(1))
        except (TypeError, ValueError):
            continue
        if not is_plausible_engine_displacement(displacement):
            continue

        variant_token = match.group(2).strip("- ")
        variant_norm = variant_token.lower()
        if variant_norm in generic_variant_words:
            continue
        if re.fullmatch(r"v\d|i\d|l\d", variant_norm, re.I):
            continue

        label = f"{displacement:.1f}L {format_named_engine_variant(variant_token)}"
        key = compact_vehicle_label(label)
        if key and key not in seen:
            variants.append(label)
            seen.add(key)

    return variants


def match_application_to_named_engine(context, named_engines):
    """Map an application cell such as "Zetec engine" to a detected named engine."""
    if not context or not named_engines:
        return None

    context_compact = compact_vehicle_label(context)
    if not context_compact:
        return None

    for engine in sorted(named_engines, key=len, reverse=True):
        parts = str(engine).split(maxsplit=1)
        if len(parts) < 2:
            continue
        variant_text = parts[1]
        variant_compact = compact_vehicle_label(variant_text)
        if not variant_compact:
            continue

        match_keys = {variant_compact}
        variant_words = [
            compact_vehicle_label(part)
            for part in re.split(r"[-\s]+", variant_text)
            if len(compact_vehicle_label(part)) >= 3
        ]
        match_keys.update(variant_words)
        for word in variant_words:
            if len(word) >= 4:
                match_keys.add(word[1:])

        if any(key and key in context_compact for key in match_keys):
            return engine

    return None


def extract_refill_application_engine_oil_capacities(text):
    """Parse refill-capacity tables with engine/application rows and oil quantities."""
    if not text:
        return {}

    text_lower = str(text).lower()
    if "refill capacities" not in text_lower or "engine oil" not in text_lower:
        return {}

    named_engines = extract_named_engine_variants_from_text(text)
    lines = [re.sub(r"\s+", " ", line).strip() for line in str(text).splitlines()]
    candidates = []

    for idx, line in enumerate(lines):
        line_lower = line.lower()
        if not line_lower:
            continue

        if "engine oil" not in line_lower and not (
            "engine" in line_lower and re.search(CAPACITY_PATTERN, line_lower, re.I)
        ):
            continue

        local_lines = [
            lines[probe_idx]
            for probe_idx in range(max(0, idx - 2), min(len(lines), idx + 3))
            if lines[probe_idx]
        ]
        local_context = " ".join(local_lines)
        local_lower = local_context.lower()

        if any(term in local_lower for term in [
            "engine coolant",
            "power steering",
            "fuel tank",
            "transaxle fluid",
            "windshield washer",
            "brake fluid",
        ]) and "engine oil" not in line_lower:
            continue

        capacity_matches = [
            match for match in re.finditer(CAPACITY_PATTERN, local_lower, re.I)
            if is_real_capacity_match(local_lower, match)
        ]
        if not capacity_matches:
            continue

        target_engine = match_application_to_named_engine(local_context, named_engines)
        if not target_engine:
            continue

        selected = choose_preferred_capacity_match(capacity_matches)
        if not selected:
            continue

        try:
            capacity = build_capacity_record_from_matches(selected, capacity_matches)
        except (ValueError, TypeError, AttributeError):
            continue

        q = capacity.get("quarts")
        if not (ENGINE_OIL_CAPACITY_MIN_QT <= q <= ENGINE_OIL_CAPACITY_MAX_QT):
            continue

        candidates.append({
            "engine": target_engine,
            "field": "with_filter",
            "capacity": capacity,
            "score": score_capacity_candidate(
                "engine oil with filter " + local_context,
                target_field="with_filter",
                engine_key=target_engine,
            ) + 12,
        })

    return select_best_capacity_candidates(candidates)


def ocr_row_has_engine_oil_filter_label(text):
    """Return True for OCR variants of an "Engine Oil with Filter" row label."""
    normalized = re.sub(r"[^a-z0-9]+", " ", str(text).lower())
    has_engine = "engine" in normalized
    has_oil = bool(re.search(r"\b(?:oil|ol|gil|o1l|0il)\b", normalized, re.I))
    has_filter = bool(re.search(r"\b(?:filter|filler|fitter|fier|fiiter|flter)\b", normalized, re.I))
    return has_engine and has_oil and has_filter


def extract_image_table_engine_oil_capacity(doc):
    """
    Recover shared oil capacity from image-backed capacity tables.

    Some manuals have a readable heading but the table cells are effectively
    image text. We detect the "Engine Oil with Filter" row using OCR word
    boxes, then OCR only the numeric columns on that same row.
    """
    if not doc:
        return {}

    target_pages = set()
    for page_num, page in enumerate(doc):
        page_text = page.get_text("text")
        page_lower = page_text.lower()
        page_lines = [re.sub(r"\s+", " ", line).strip().lower() for line in page_text.splitlines()]
        if not page.get_images():
            continue

        has_spec_heading = any(
            line in {"capacities and specifications", "engine specifications", "engine specification"}
            for line in page_lines
        )
        has_capacity_table_text = (
            "capacities and specifications" in page_lower
            and "capacities" in page_lower
        )
        has_dotted_index_entry = bool(re.search(
            r"(?:capacities and specifications|engine specifications)\s*\.{3,}",
            page_lower,
            re.I,
        ))

        if (has_spec_heading or has_capacity_table_text) and not has_dotted_index_entry:
            target_pages.add(page_num)

    candidates = []
    for page_num in sorted(target_pages):
        try:
            page = doc[page_num]
            pix = page.get_pixmap(matrix=fitz.Matrix(8, 8), colorspace=fitz.csRGB)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            gray = ImageOps.grayscale(img)
            variants = [
                gray,
                gray.filter(ImageFilter.SHARPEN),
                gray.point(lambda x: 0 if x < 190 else 255),
                gray.point(lambda x: 0 if x < 210 else 255),
            ]

            row_boxes = []
            for variant in variants:
                data = pytesseract.image_to_data(
                    variant,
                    config="--psm 6",
                    output_type=pytesseract.Output.DICT,
                )
                grouped_rows = {}
                for idx, word in enumerate(data.get("text", [])):
                    word = str(word).strip()
                    if not word:
                        continue
                    key = (
                        data["block_num"][idx],
                        data["par_num"][idx],
                        data["line_num"][idx],
                    )
                    grouped_rows.setdefault(key, []).append({
                        "text": word,
                        "left": data["left"][idx],
                        "top": data["top"][idx],
                        "width": data["width"][idx],
                        "height": data["height"][idx],
                    })

                for row in grouped_rows.values():
                    row_text = " ".join(item["text"] for item in sorted(row, key=lambda item: item["left"]))
                    if not ocr_row_has_engine_oil_filter_label(row_text):
                        continue
                    left = min(item["left"] for item in row)
                    top = min(item["top"] for item in row)
                    right = max(item["left"] + item["width"] for item in row)
                    bottom = max(item["top"] + item["height"] for item in row)
                    box = (left, top, right, bottom)
                    if box not in row_boxes:
                        row_boxes.append(box)

            for left, top, right, bottom in row_boxes:
                row_height = max(40, bottom - top)
                center_y = (top + bottom) // 2
                crop_left = max(int(pix.width * 0.55), right + 20)
                crop_right = int(pix.width * 0.98)
                crop_top = max(0, center_y - max(45, int(row_height * 0.9)))
                crop_bottom = min(pix.height, center_y + max(55, int(row_height * 1.1)))
                if crop_right <= crop_left or crop_bottom <= crop_top:
                    continue

                crop = gray.crop((crop_left, crop_top, crop_right, crop_bottom))
                crop_variants = [
                    crop,
                    crop.filter(ImageFilter.SHARPEN),
                    crop.point(lambda x: 0 if x < 190 else 255),
                    crop.point(lambda x: 0 if x < 210 else 255),
                ]
                for crop_img in crop_variants:
                    for config in ("--psm 6", "--psm 7", "--psm 11", "--psm 13"):
                        text = pytesseract.image_to_string(crop_img, config=config)
                        capacity = capacity_from_ocr_numeric_text(text)
                        if not capacity:
                            continue
                        q = capacity.get("quarts")
                        if not (ENGINE_OIL_CAPACITY_MIN_QT <= q <= ENGINE_OIL_CAPACITY_MAX_QT):
                            continue
                        candidates.append({
                            "engine": "unknown_engine",
                            "field": "with_filter",
                            "capacity": capacity,
                            "score": score_capacity_candidate(
                                "engine oil with filter " + text,
                                target_field="with_filter",
                                engine_key="unknown_engine",
                            ) + 14,
                        })
        except Exception as e:
            print(f"      Warning: image table oil-capacity OCR failed on page {page_num}: {str(e)}")
            continue

    return select_best_capacity_candidates(candidates)


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


def get_local_pdfs(folder_path=LOCAL_MANUALS_FOLDER):
    """Recursively collect PDF files under the local Manuals folder."""
    manuals_path = Path(folder_path)
    if not manuals_path.exists():
        print(f"Local manuals folder not found: {manuals_path.resolve()}")
        return []

    pdfs = []
    for pdf_path in sorted(manuals_path.rglob("*.pdf")):
        pdfs.append({
            "name": pdf_path.name,
            "path": pdf_path
        })

    return pdfs


def load_local_pdf(file_path):
    """Load a local PDF into an in-memory byte buffer."""
    buffer = io.BytesIO(Path(file_path).read_bytes())
    buffer.seek(0)
    return buffer


def choose_pdf_source():
    """Ask which PDF source should feed the extraction run."""
    print("\nChoose PDF extraction source:")
    print("1. Google Drive")
    print(f"2. Local {LOCAL_MANUALS_FOLDER} folder")

    choice = input("Enter choice (1 or 2): ").strip()

    match choice:
        case "1":
            return "drive"
        case "2":
            return "local"
        case _:
            print("Invalid choice. Using Google Drive.")
            return "drive"


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


def normalize_ocr_oil_text(text):
    """Repair common OCR confusions in oil grades before regex extraction."""
    if not text:
        return text

    normalized = str(text)
    normalized = re.sub(r"\bSAL\b", "SAE", normalized, flags=re.I)

    # Common OCR issues: OW-30 -> 0W-30, 1OW-30 -> 10W-30, 2OW-50 -> 20W-50.
    normalized = re.sub(
        r"\b([012])O(?=W[-\u2013\u2014]?\s*(?:16|20|30|40|50|60)\b)",
        lambda m: m.group(1) + "0",
        normalized,
        flags=re.I,
    )
    normalized = re.sub(
        r"\bO(?=W[-\u2013\u2014]?\s*(?:16|20|30|40|50|60)\b)",
        "0",
        normalized,
        flags=re.I,
    )
    normalized = re.sub(
        r"\bS(?=W[-\u2013\u2014]?\s*(?:16|20|30|40|50|60)\b)",
        "5",
        normalized,
        flags=re.I,
    )

    return normalized


def to_quarts_liters(value, unit):
    """Convert one capacity value to both quarts and liters."""
    value = float(value)
    unit_lower = unit.lower().rstrip(".")
    if unit_lower.startswith("l"):
        return round(value / 0.946352946, 2), value
    if unit_lower.startswith("gal"):
        quarts = value * 4.0
        return round(quarts, 2), round(value * 3.785411784, 1)
    return value, round(value * 0.946352946, 1)



def parse_filename(name):
    """Extract year, make, and model from flexible manual filenames."""
    clean_name = name.replace("Copy of ", "").strip()
    stem = re.sub(r"\.pdf$", "", clean_name, flags=re.I)
    if is_generic_manual_filename(stem):
        return None, None, None

    match = re.match(r"(\d{4})-([^-]+)-(.+)$", stem, re.I)
    if match:
        year, make, model = match.groups()
        model = model.replace("-OM", "").replace("-UG", "").replace("-UM", "")
        return int(year), make.capitalize(), model.capitalize()

    year = None
    year_match = re.search(r"(?<!\d)((?:19|20)\d{2})(?!\d)", stem)
    if year_match:
        year = int(year_match.group())

    reference = load_vehicle_reference()
    normalized_stem = normalize_vehicle_label(stem)
    compact_stem = compact_vehicle_label(stem)

    make = None
    model = None

    model_candidates = []
    for ref_make, ref_models in reference.items():
        make_norm = normalize_vehicle_label(ref_make)
        make_compact = compact_vehicle_label(ref_make)
        make_tokens = [
            token for token in make_norm.split()
            if len(token) >= 4 and token not in {"benz", "auto", "cars"}
        ]
        make_in_name = bool(
            make_norm and make_norm in normalized_stem
            or make_compact and make_compact in compact_stem
            or any(token in normalized_stem.split() for token in make_tokens)
        )

        for ref_model in ref_models:
            model_norm = normalize_vehicle_label(ref_model)
            model_compact = compact_vehicle_label(ref_model)
            if not model_norm:
                continue
            if filename_contains_vehicle_model(stem, ref_model):
                model_candidates.append((ref_make, ref_model, len(model_compact or model_norm), make_in_name))

    if model_candidates and any(item[3] for item in model_candidates):
        model_candidates.sort(key=lambda item: (item[3], item[2]), reverse=True)
        make, model, _, _ = model_candidates[0]
    elif model_candidates and not is_generic_manual_filename(stem):
        model_candidates.sort(key=lambda item: item[2], reverse=True)
        best_length = model_candidates[0][2]
        best_candidates = [item for item in model_candidates if item[2] == best_length]
        best_pairs = {(item[0], item[1]) for item in best_candidates}
        if len(best_pairs) == 1:
            make, model, _, _ = best_candidates[0]
    else:
        for ref_make in reference.keys():
            make_norm = normalize_vehicle_label(ref_make)
            make_compact = compact_vehicle_label(ref_make)
            if make_norm in normalized_stem or (make_compact and make_compact in compact_stem):
                make = ref_make
                break

    return year, make, model


def build_multi_engine_data(engine_caps, oil_scores, oil_temps, engine_oil_map):
    """Build the final per-engine capacity and oil recommendation records."""
    engine_data = {}

    for eng, cap in engine_caps.items():
        with_filter = cap.get("with_filter")
        without_filter = cap.get("without_filter")
        oil_list = []

        if oil_scores:
            valid_oils = engine_oil_map.get(eng, [])
            oils_from_engine_map = bool(valid_oils)
            
            if not valid_oils:
                base_eng_size = eng.split()[0]
                valid_oils = engine_oil_map.get(base_eng_size, [])
                oils_from_engine_map = bool(valid_oils)
            
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
                def engine_oil_rank(oil):
                    """Rank candidate oils by engine mapping, score, and broad-temperature usability."""
                    score = oil_scores.get(oil, 0)
                    temps = oil_temps.get(oil, set())
                    map_order_bonus = 0
                    if oils_from_engine_map and oil in valid_oils:
                        map_order_bonus = len(valid_oils) - valid_oils.index(oil)
                    cold_only = any(
                        isinstance(t, str) and (
                            "cold weather" in t.lower()
                            or "-20f" in t.lower()
                            or "-22f" in t.lower()
                            or "below" in t.lower()
                        )
                        for t in temps
                    ) and not any(
                        isinstance(t, str) and "all temperatures" in t.lower()
                        for t in temps
                    )
                    zero_weight = oil.startswith("0W-")
                    return (map_order_bonus, score, not cold_only, not zero_weight)

                engine_best = max(valid_oils, key=engine_oil_rank)
                engine_max_score = oil_scores.get(engine_best, 0)
            else:
                engine_best = max(oil_scores, key=oil_scores.get)
                engine_max_score = oil_scores[engine_best]
                valid_oils = list(oil_scores.keys())

            for oil in valid_oils:
                score = oil_scores.get(oil, 0)
                temps = sorted(get_temperature_with_fallback(oil_temps.get(oil, []), oil))
                
                has_actual_temps = any(
                    ("F" in t or "C" in t or 
                     "above" in t or "below" in t or "range:" in t or 
                     "weather" in t or "temperatures" in t or "cold" in t or "hot" in t)
                    for t in temps
                )

                if oils_from_engine_map:
                    oil_list.append({
                        "oil_type": oil,
                        "recommendation_level": "primary" if oil == engine_best else "secondary",
                        "temperature_condition": list(temps),
                    })
                elif score >= engine_max_score - 1 or has_actual_temps:
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


def build_oil_only_engine_data(all_engines, oil_scores, oil_temps):
    """Create engine records when only oil recommendations are known."""
    if not oil_scores:
        return {}

    primary = max(oil_scores, key=oil_scores.get)
    max_score = oil_scores[primary]
    oil_list = []
    for oil, score in oil_scores.items():
        temps = sorted(get_temperature_with_fallback(oil_temps.get(oil, []), oil))
        has_actual_temps = any(
            ("F" in t or "C" in t or "above" in t or "below" in t or "range:" in t or
             "weather" in t or "temperatures" in t or "cold" in t or "hot" in t)
            for t in temps
        )
        if score >= max_score - 2 or has_actual_temps:
            oil_list.append({
                "oil_type": oil,
                "recommendation_level": "primary" if oil == primary else "secondary",
                "temperature_condition": temps,
            })

    if not oil_list:
        return {}

    engine_keys = all_engines or ["unknown_engine"]
    return {
        eng: {
            "oil_capacity": {"with_filter": None, "without_filter": None},
            "oil_recommendations": [dict(item) for item in oil_list],
        }
        for eng in engine_keys
    }


def expand_shared_capacity_to_detected_engines(engine_caps, all_engines):
    """
    Copy a shared unknown-engine capacity to detected engines when the PDF gives no
    per-engine rows.
    """
    if not engine_caps or "unknown_engine" not in engine_caps or not all_engines:
        return engine_caps

    shared_cap = engine_caps.pop("unknown_engine")
    existing_bases = {engine_identity_key(eng) for eng in engine_caps.keys() if eng != "unknown_engine"}

    for eng in all_engines:
        eng_base = engine_identity_key(eng)
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
    variant_by_bodyless_base = {}
    variant_by_displacement = {}
    for eng in all_engines:
        base = engine_identity_key(eng)
        bodyless_base = engine_identity_key(eng, strip_body_style=True)
        if len(eng.split()) > 1:
            variant_by_base[base] = eng
            variant_by_bodyless_base[bodyless_base] = eng
            displacement = get_engine_displacement(eng)
            if is_plausible_engine_displacement(displacement):
                variant_by_displacement[round(displacement, 1)] = eng

    aligned = {}
    for eng, cap in engine_caps.items():
        if eng == "unknown_engine":
            aligned[eng] = cap
            continue

        base = engine_identity_key(eng)
        bodyless_base = engine_identity_key(eng, strip_body_style=True)
        displacement = get_engine_displacement(eng)
        displacement_key = round(displacement, 1) if is_plausible_engine_displacement(displacement) else None
        aligned_key = (
            variant_by_base.get(base)
            or variant_by_bodyless_base.get(bodyless_base)
            or (
                variant_by_displacement.get(displacement_key)
                if displacement_key is not None and not engine_label_has_type(eng)
                else None
            )
            or eng
        )
        aligned[aligned_key] = cap

    return aligned


def filter_engine_caps_to_detected_engines(engine_caps, all_engines):
    """Drop capacity keys that do not match trusted detected engine bases."""
    if not engine_caps or not all_engines:
        return engine_caps

    valid_bases = {engine_identity_key(eng) for eng in all_engines}
    valid_bodyless_bases = {engine_identity_key(eng, strip_body_style=True) for eng in all_engines}
    valid_displacements = {
        round(displacement, 1)
        for displacement in (get_engine_displacement(eng) for eng in all_engines)
        if is_plausible_engine_displacement(displacement)
    }
    filtered = {}
    unknown_cap = engine_caps.get("unknown_engine")

    for eng, cap in engine_caps.items():
        if eng == "unknown_engine":
            continue

        displacement = get_engine_displacement(eng)
        if (
            engine_identity_key(eng) in valid_bases
            or engine_identity_key(eng, strip_body_style=True) in valid_bodyless_bases
            or (
                is_plausible_engine_displacement(displacement)
                and round(displacement, 1) in valid_displacements
                and not engine_label_has_type(eng)
            )
        ):
            filtered[eng] = cap

    if filtered:
        return filtered
    if unknown_cap:
        return {"unknown_engine": unknown_cap}
    return engine_caps


def add_capacity_backed_engine_candidates(all_engines, engine_caps, candidate_engines):
    """
    Add engines missed by the preferred detector only when an engine-oil
    capacity row independently names the same engine.
    """
    if not engine_caps or not candidate_engines:
        return all_engines

    all_engines = list(all_engines or [])
    known_bases = {engine_identity_key(eng) for eng in all_engines}
    known_bodyless_bases = {engine_identity_key(eng, strip_body_style=True) for eng in all_engines}
    known_displacements = {
        round(displacement, 1)
        for displacement in (get_engine_displacement(eng) for eng in all_engines)
        if is_plausible_engine_displacement(displacement)
    }

    cap_bases = set()
    cap_bodyless_bases = set()
    cap_displacements = set()
    for cap_eng in engine_caps:
        if cap_eng == "unknown_engine":
            continue
        cap_bases.add(engine_identity_key(cap_eng))
        cap_bodyless_bases.add(engine_identity_key(cap_eng, strip_body_style=True))
        displacement = get_engine_displacement(cap_eng)
        if is_plausible_engine_displacement(displacement):
            cap_displacements.add(round(displacement, 1))

    for eng in candidate_engines:
        eng_base = engine_identity_key(eng)
        eng_bodyless_base = engine_identity_key(eng, strip_body_style=True)
        displacement = get_engine_displacement(eng)
        displacement_key = round(displacement, 1) if is_plausible_engine_displacement(displacement) else None

        if (
            eng_base in known_bases
            or eng_bodyless_base in known_bodyless_bases
            or (displacement_key is not None and displacement_key in known_displacements)
        ):
            continue

        if (
            eng_base in cap_bases
            or eng_bodyless_base in cap_bodyless_bases
            or (displacement_key is not None and displacement_key in cap_displacements)
        ):
            all_engines.append(eng)
            known_bases.add(eng_base)
            known_bodyless_bases.add(eng_bodyless_base)
            if displacement_key is not None:
                known_displacements.add(displacement_key)

    return all_engines


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
    if is_layout_engine_type(engine_label):
        return True

    parts = str(engine_label).split()
    if len(parts) <= 1:
        return False

    trailing = " ".join(parts[1:])
    return bool(re.search(ENGINE_TYPE_PATTERN, trailing, re.I))


def engine_label_has_layout(engine_label):
    """Return True when an engine key already includes a layout token such as I4 or V6."""
    if not engine_label or engine_label == "unknown_engine":
        return True
    if is_layout_engine_type(engine_label):
        return True

    parts = str(engine_label).split()
    if len(parts) <= 1:
        return False

    trailing = " ".join(parts[1:])
    return any(is_layout_engine_type(token) for token in trailing.split())


def select_single_layout_type(engine_types):
    """Pick one vehicle-wide layout type only when it is unambiguous."""
    layout_types = []
    for engine_type in engine_types or []:
        normalized = normalize_engine_type_token(engine_type)
        if re.fullmatch(r"(?:I|V|W|H|F)\d{1,2}", normalized) or normalized.startswith("FLAT") or normalized in {"BOXER", "ROTARY", "WANKEL"}:
            if normalized not in layout_types:
                layout_types.append(normalized)

    return layout_types[0] if len(layout_types) == 1 else None


def is_layout_engine_type(engine_type):
    """Return True for cylinder/layout types such as I4, V6, W16, or BOXER."""
    normalized = normalize_engine_type_token(engine_type)
    return bool(
        re.fullmatch(r"(?:I|V|W|H|F)\d{1,2}", normalized)
        or normalized.startswith("FLAT")
        or normalized in {"BOXER", "ROTARY", "WANKEL"}
    )


def extract_layout_engine_tokens(text):
    """Return layout-only engine tokens from spec-table cells, such as V6."""
    if not text:
        return []

    tokens = []
    for match in re.finditer(
        r"\b(?:v\s*-?\s*(?:3|4|5|6|8|10|12|16|20|24)|i\s*-?\s*[3-8]|l\s*-?\s*[3-8]|h\s*-?\s*4|w\s*-?\s*(?:8|12|16)|f\s*-?\s*8|inline\s*-?\s*[3-8]|flat\s*-?\s*(?:3|4|6|8|12)|boxer|rotary|wankel)\b",
        str(text),
        re.I,
    ):
        token = normalize_engine_type_token(match.group())
        if token and is_layout_engine_type(token) and token not in tokens:
            tokens.append(token)

    for match in re.finditer(CYLINDER_PATTERN, str(text), re.I):
        token = normalize_cylinder_match(match)
        if token and token not in tokens:
            tokens.append(token)

    return tokens


def filter_engine_types_by_detected_engines(engine_types, all_engines):
    """Keep only layout types supported by the detected engine labels when available."""
    if not engine_types or not all_engines:
        return engine_types

    explicit_layouts = []
    for engine_label in all_engines:
        token_parts = str(engine_label).split()[1:]
        if not token_parts and is_layout_engine_type(engine_label):
            token_parts = [str(engine_label)]
        for token in token_parts:
            normalized = normalize_engine_type_token(token)
            if is_layout_engine_type(normalized) and normalized not in explicit_layouts:
                explicit_layouts.append(normalized)

    if not explicit_layouts:
        return engine_types

    filtered = []
    for engine_type in engine_types:
        normalized = normalize_engine_type_token(engine_type)
        if is_layout_engine_type(normalized) and normalized not in explicit_layouts:
            continue
        filtered.append(engine_type)

    return filtered


def add_missing_engine_type_to_keys(engine_data, engine_types):
    """
    Append an unambiguous vehicle-wide layout to bare displacement keys.
    Example: {"1.4L": ...} with engine_types ["I4"] becomes {"1.4L I4": ...}.
    """
    layout_type = select_single_layout_type(engine_types)
    if not engine_data or not layout_type:
        return engine_data

    distinct_displacements = {
        round(disp, 1)
        for disp in (get_engine_displacement(label) for label in engine_data)
        if is_plausible_engine_displacement(disp)
    }
    if len(distinct_displacements) != 1:
        return engine_data

    relabeled = {}
    for engine_label, engine_info in engine_data.items():
        if is_layout_engine_type(engine_label):
            relabeled[engine_label] = engine_info
            continue
        displacement = get_engine_displacement(engine_label)
        if not is_plausible_engine_displacement(displacement):
            relabeled[engine_label] = engine_info
            continue
        if engine_label_has_layout(engine_label):
            relabeled[engine_label] = engine_info
            continue
        if engine_label_has_type(engine_label):
            relabeled[f"{engine_label} {layout_type}"] = engine_info
            continue

        relabeled[f"{engine_label} {layout_type}"] = engine_info

    return relabeled


def build_capacity_record(value, unit):
    """Create a normalized capacity record with both U.S. quarts and liters."""
    q, l = to_quarts_liters(value, unit)
    return {"quarts": q, "liters": l}


def build_capacity_record_from_matches(selected_match, capacity_matches):
    """Use the selected U.S. amount plus the manual's metric conversion when present."""
    capacity = build_capacity_record(selected_match.group(1), selected_match.group(2))
    selected_unit = selected_match.group(2).lower()

    if "qt" in selected_unit or "quart" in selected_unit or "gal" in selected_unit:
        for match in capacity_matches:
            unit = match.group(2).lower()
            if unit.startswith("l"):
                try:
                    capacity["liters"] = float(match.group(1))
                except (ValueError, TypeError):
                    pass
                break

    return capacity


def choose_preferred_capacity_match(capacity_matches):
    """Prefer the original U.S. capacity value over parenthesized metric conversions."""
    if not capacity_matches:
        return None

    for match in capacity_matches:
        unit = match.group(2).lower()
        if "qt" in unit or "quart" in unit or "gal" in unit:
            return match

    return capacity_matches[0]


def detect_capacity_field(text):
    """Classify a capacity line as with-filter or without-filter when possible."""
    text_lower = str(text).lower()
    if any(label in text_lower for label in WITH_FILTER_LABELS):
        return "with_filter"
    if any(label in text_lower for label in WITHOUT_FILTER_LABELS):
        return "without_filter"
    return None


def score_capacity_candidate(text, target_field=None, engine_key=None):
    """Score a capacity candidate from surrounding context before final selection."""
    text_lower = str(text).lower()
    score = 0

    if any(term in text_lower for term in ENGINE_OIL_STOP_TERMS):
        score -= 12
    if "adding engine oil" in text_lower or "dipstick" in text_lower:
        score -= 8
    if any(term in text_lower for term in CAPACITY_REFILL_CONTEXT):
        score -= 14
    if any(term in text_lower for term in NON_CAPACITY_OIL_QUANTITY_CONTEXT):
        score -= 18
    if "oil filter" in text_lower and "engine oil" not in text_lower and "with filter" not in text_lower:
        score -= 5
    if "for specific engine oil capacities" in text_lower:
        score -= 16

    if "engine oil" in text_lower:
        score += 8
    if "oil capacity" in text_lower or "engine oil capacity" in text_lower:
        score += 6
    if "crankcase" in text_lower:
        score += 4
    if "specifications" in text_lower or "capacities" in text_lower:
        score += 2
    if "with filter" in text_lower or "with oil filter" in text_lower or "including filter" in text_lower:
        score += 4
    if "without filter" in text_lower or "excluding the oil filter" in text_lower:
        score += 4
    if re.search(ENGINE_PATTERN, text_lower, re.I):
        score += 2
    if re.search(r"\d+\.?\d*\s*(?:quarts?|qts?|qt\.?|gallons?|gal\.?).{0,12}\d+\.?\d*\s*(?:liters?|litres?|l\b)", text_lower, re.I):
        score += 3

    detected_field = detect_capacity_field(text_lower)
    if detected_field and target_field and detected_field == target_field:
        score += 5
    elif detected_field and target_field and detected_field != target_field:
        score -= 4

    if engine_key and engine_key != "unknown_engine":
        score += 2

    capacity_matches = [
        match for match in re.finditer(CAPACITY_PATTERN, text_lower, re.I)
        if is_real_capacity_match(text_lower, match)
    ]
    if capacity_matches:
        selected = choose_preferred_capacity_match(capacity_matches)
        if selected:
            try:
                q, _ = to_quarts_liters(selected.group(1), selected.group(2))
            except (TypeError, ValueError, AttributeError):
                q = None
            if q is not None and q <= 1.5 and engine_key == "unknown_engine":
                score -= 18

    return score


def extract_wildcard_oil_candidates(text):
    """Expand viscosity-class patterns like SAE 0W-X where X stands for 30, 40 or 50."""
    if not text:
        return []

    candidates = []
    text_lower = str(text).lower()
    windows = re.finditer(
        r"sae\s+\d+\s*w\s*-\s*x.{0,220}?stands\s+for.{0,120}",
        text_lower,
        re.I,
    )

    for match in windows:
        window = text_lower[match.start():min(len(text_lower), match.end() + 80)]
        base_weights = re.findall(r"sae\s+((?:0|5|10|15|20|25)w)\s*-\s*x", window, re.I)
        grades = re.findall(r"\b(16|20|30|40|50|60)\b", window)
        if not base_weights or not grades:
            continue

        seen = set()
        for base in base_weights:
            for grade in grades:
                oil = normalize_oil(f"{base.upper()}-{grade}")
                if oil in seen:
                    continue
                seen.add(oil)
                candidates.append({
                    "oil": oil,
                    "window": window,
                    "base": base.upper(),
                })

    return candidates


def extract_listed_oil_candidates(text):
    """Extract explicit SAE lists such as 'SAE 0W-30, SAE 5W-30 or SAE 5W-40'."""
    if not text:
        return []

    candidates = []
    text_lower = str(text).lower()
    trigger_pattern = re.compile(
        r"(viscosity grade of|viscosity grades|use only engine oils of viscosity class|engine oil[s]?(?: of)? viscosity class)",
        re.I,
    )

    for match in trigger_pattern.finditer(text_lower):
        window_start = max(0, match.start() - 160)
        window_end = min(len(text), match.end() + 260)
        window = text[window_start:window_end]
        oils = extract_oil_types_from_text(window)

        if oils:
            candidates.append({"window": window, "oils": oils})

    return candidates


def select_best_capacity_candidates(candidates):
    """Keep the strongest capacity candidate for each engine and filter field."""
    selected = {}
    for candidate in candidates:
        if candidate.get("score", 0) < 1:
            continue
        engine_key = candidate["engine"]
        field = candidate["field"]
        current = selected.setdefault(engine_key, {"with_filter": None, "without_filter": None})
        existing = current.get(field)
        if existing is None or candidate["score"] > existing["score"]:
            current[field] = candidate

    result = {}
    for engine_key, fields in selected.items():
        with_filter = fields.get("with_filter")
        without_filter = fields.get("without_filter")
        if not with_filter and not without_filter:
            continue
        if with_filter and without_filter:
            with_q = (with_filter.get("capacity") or {}).get("quarts")
            without_q = (without_filter.get("capacity") or {}).get("quarts")
            if (
                isinstance(with_q, (int, float))
                and isinstance(without_q, (int, float))
                and without_q >= with_q
            ):
                without_filter = None
        result[engine_key] = {
            "with_filter": dict(with_filter["capacity"]) if with_filter else None,
            "without_filter": dict(without_filter["capacity"]) if without_filter else None,
        }
    return result


def extract_ordered_engine_oil_capacity_from_text(text):
    """
    Recover oil capacity from OCR tables where labels and numeric columns were
    split apart but row order is preserved around Cooling System -> Engine Oil.
    """
    if not text:
        return None

    text_lower = str(text).lower()
    if "engine oil with filter" not in text_lower:
        return None

    for cooling_match in re.finditer(r"cooling\s+system", text_lower, re.I):
        window = text_lower[cooling_match.start():min(len(text_lower), cooling_match.end() + 500)]
        if (
            "engine oil with filter" not in window
            and "fuel capacity" not in window
            and len(re.findall(CAPACITY_PATTERN, window, re.I)) < 2
        ):
            continue

        quart_matches = []
        for match in re.finditer(CAPACITY_PATTERN, window, re.I):
            if not is_real_capacity_match(window, match):
                continue
            try:
                q, l = to_quarts_liters(match.group(1), match.group(2))
            except (TypeError, ValueError, AttributeError):
                continue
            if ENGINE_OIL_CAPACITY_MIN_QT <= q <= ENGINE_OIL_CAPACITY_MAX_QT:
                quart_matches.append((q, l))

        if len(quart_matches) >= 2:
            cooling_q = quart_matches[0][0]
            oil_q, oil_l = quart_matches[1]
            if cooling_q > oil_q and oil_q <= 8.0:
                return {"quarts": oil_q, "liters": round(oil_q * 0.946352946, 1)}

    return None


def has_ambiguous_model_oil_capacity_table(lines):
    """Detect multi-model oil-capacity tables that should not collapse to one shared value."""
    if not lines:
        return False

    for idx, line in enumerate(lines):
        line_lower = str(line).lower()
        if (
            "engine oil filling quantity" not in line_lower
            and "engine oil quality and filling quantity" not in line_lower
        ):
            continue

        window_lines = [str(x).strip() for x in lines[idx:min(len(lines), idx + 30)] if str(x).strip()]
        window_text = " ".join(window_lines).lower()
        if "model" not in window_text:
            continue

        quart_values = set()
        for match in re.finditer(CAPACITY_PATTERN, window_text, re.I):
            if not is_real_capacity_match(window_text, match):
                continue
            try:
                quarts, _ = to_quarts_liters(match.group(1), match.group(2))
            except (TypeError, ValueError, AttributeError):
                continue
            quart_values.add(round(quarts, 2))

        if len(quart_values) >= 2:
            return True

    return False


def has_external_engine_oil_capacity_reference(text):
    """Return True when the manual delegates oil capacity values to an external source."""
    text_lower = str(text).lower()
    return (
        "for specific engine oil capacities" in text_lower
        or "see the most current information" in text_lower and "engine oil" in text_lower
    )


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
_engine_family_reference_cache = None


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
            if isinstance(make, str) and make.startswith("_"):
                continue
            if not isinstance(make, str) or not isinstance(models, list):
                continue
            clean_models = [m for m in models if isinstance(m, str) and m.strip()]
            if clean_models:
                normalized[make.strip()] = clean_models

    _vehicle_reference_cache = normalized
    return _vehicle_reference_cache


def load_engine_family_reference():
    """Load optional engine family/technology labels from vehicle_reference.json."""
    global _engine_family_reference_cache
    if _engine_family_reference_cache is not None:
        return _engine_family_reference_cache

    ref_path = Path(__file__).with_name(REFERENCE_FILE)
    if not ref_path.exists():
        _engine_family_reference_cache = []
        return _engine_family_reference_cache

    try:
        with ref_path.open("r", encoding="utf-8") as f:
            data = json.load(f)
    except (OSError, json.JSONDecodeError):
        data = {}

    families = []
    if isinstance(data, dict):
        raw_families = data.get("_engine_families", [])
        if isinstance(raw_families, list):
            for family in raw_families:
                if isinstance(family, str) and family.strip():
                    families.append(family.strip())

    _engine_family_reference_cache = families
    return _engine_family_reference_cache


def normalize_vehicle_label(value):
    """Normalize make/model text for case-insensitive reference matching."""
    if not value:
        return ""

    text = value.lower()
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def compact_vehicle_label(value):
    """Compact a vehicle label so filename tokens like 7Series match 7 Series."""
    return re.sub(r"[^a-z0-9]+", "", str(value).lower())


def vehicle_filename_tokens(stem):
    """Return normalized filename tokens used for make/model matching."""
    return [
        token
        for token in re.split(r"[^a-z0-9]+", str(stem).lower())
        if token
    ]


def filename_contains_vehicle_model(stem, model):
    """Match a model in a filename without letting years create false hits."""
    model_norm = normalize_vehicle_label(model)
    model_compact = compact_vehicle_label(model)
    if not model_norm or not model_compact:
        return False

    normalized_stem = normalize_vehicle_label(stem)
    tokens = vehicle_filename_tokens(stem)

    # Phrase match covers normal separators: "Ford-Focus", "7 Series".
    if re.search(r"\b" + re.escape(model_norm).replace(r"\ ", r"\s+") + r"\b", normalized_stem, re.I):
        return True

    # Compact token match covers filename forms such as "7Series" and "x3".
    # Avoid pure short numeric models matching the year or other incidental numbers.
    if model_compact in tokens:
        return not (model_compact.isdigit() and len(model_compact) < 3)

    # Allow compact matching across adjacent non-year tokens, but keep it anchored
    # to token boundaries so "2" does not match the "2013" year.
    non_year_tokens = [
        token for token in tokens
        if not re.fullmatch(r"(?:19|20)\d{2}", token)
    ]
    for start in range(len(non_year_tokens)):
        joined = ""
        for token in non_year_tokens[start:start + 4]:
            joined += token
            if joined == model_compact:
                return not (model_compact.isdigit() and len(model_compact) < 3)
            if len(joined) >= len(model_compact):
                break

    return False


def strip_parenthetical_body_style(value):
    """Remove body-style suffixes like '(SUV)' to improve variant matching."""
    if not value:
        return ""
    stripped = re.sub(r"\((?:suv|coupe|coup.?|sedan|wagon|hatchback|convertible|roadster|cabriolet)[^)]*\)", "", str(value), flags=re.I)
    return re.sub(r"\s+", " ", stripped).strip(" ,:-")


def canonicalize_engine_variant_label(value, model=None):
    """Normalize model-based variant labels so table rows match detected engine keys."""
    if not value:
        return ""

    cleaned = str(value)
    cleaned = cleaned.replace("\u2010", "-").replace("\u2011", "-").replace("\u2013", "-").replace("\u2014", "-")
    cleaned = cleaned.replace("4MATIC +", "4MATIC+")
    cleaned = re.sub(r"\s*\+\s*", "+", cleaned)
    cleaned = re.sub(r"\bCoup.\b", "Coupe", cleaned, flags=re.I)
    cleaned = re.sub(r"\s+", " ", cleaned).strip(" ,:-")

    if model:
        model_pattern = r"\b" + re.escape(str(model)).replace(r"\ ", r"\s+") + r"\b"
        cleaned = re.sub(model_pattern, "", cleaned, count=1, flags=re.I)
        cleaned = re.sub(r"\s+", " ", cleaned).strip(" ,:-")

    return cleaned


def engine_identity_key(value, model=None, strip_body_style=False):
    """Build a stable comparison key for variant labels without collapsing distinct trims."""
    canonical = canonicalize_engine_variant_label(value, model=model)
    if strip_body_style:
        canonical = strip_parenthetical_body_style(canonical)
    return compact_vehicle_label(canonical)


def match_model_label_to_detected_engine(label, detected_engines=None, model=None):
    """Match a model row label to the closest detected engine variant when possible."""
    canonical = canonicalize_engine_variant_label(label, model=model)
    if not canonical:
        return ""
    if not detected_engines:
        return canonical

    candidate_forms = []
    for raw_value in (label, canonical, strip_parenthetical_body_style(label), strip_parenthetical_body_style(canonical)):
        normalized = canonicalize_engine_variant_label(raw_value, model=model)
        compact = compact_vehicle_label(normalized)
        if normalized and compact and compact not in {item[1] for item in candidate_forms}:
            candidate_forms.append((normalized, compact))

    engine_forms = []
    for engine in detected_engines:
        canonical_engine = canonicalize_engine_variant_label(engine, model=model)
        compact_engine = compact_vehicle_label(canonical_engine)
        compact_engine_bodyless = compact_vehicle_label(strip_parenthetical_body_style(canonical_engine))
        engine_forms.append((engine, canonical_engine, compact_engine, compact_engine_bodyless))

    for _, candidate_compact in candidate_forms:
        for engine, _, compact_engine, compact_engine_bodyless in engine_forms:
            if candidate_compact in {compact_engine, compact_engine_bodyless}:
                return engine

    ranked = []
    for _, candidate_compact in candidate_forms:
        for engine, _, compact_engine, compact_engine_bodyless in engine_forms:
            if candidate_compact and compact_engine and (candidate_compact in compact_engine or compact_engine in candidate_compact):
                ranked.append((engine, max(len(candidate_compact), len(compact_engine))))
            elif candidate_compact and compact_engine_bodyless and (
                candidate_compact in compact_engine_bodyless or compact_engine_bodyless in candidate_compact
            ):
                ranked.append((engine, max(len(candidate_compact), len(compact_engine_bodyless))))

    if ranked:
        ranked.sort(key=lambda item: item[1], reverse=True)
        return ranked[0][0]

    return canonical


def extract_oil_types_from_text(text):
    """Extract explicit oil viscosities, including shorthand such as SAE 0W-/5W-40."""
    if not text:
        return []

    normalized = normalize_ocr_oil_text(text)
    oils = []
    seen = set()

    shorthand_patterns = [
        re.compile(
            r"\b((?:0|5|10|15|20|25)W)\s*-\s*/\s*((?:0|5|10|15|20|25)W)\s*-\s*(16|20|30|40|50|60)\b",
            re.I,
        ),
        re.compile(
            r"\b((?:0|5|10|15|20|25)W)\s*/\s*((?:0|5|10|15|20|25)W)\s*-\s*(16|20|30|40|50|60)\b",
            re.I,
        ),
    ]

    for pattern in shorthand_patterns:
        for match in pattern.finditer(normalized):
            grade = match.group(3)
            for base in (match.group(1), match.group(2)):
                oil = normalize_oil(f"{base}-{grade}")
                if oil not in seen:
                    seen.add(oil)
                    oils.append(oil)

    for base, grade in re.findall(OIL_PATTERN, normalized, re.I):
        oil = normalize_oil(f"{base}W-{grade}")
        if oil not in seen:
            seen.add(oil)
            oils.append(oil)

    return oils


def cylinder_count_to_layout_token(cylinder_count):
    """Convert a plain cylinder count to a likely layout token for column tables."""
    try:
        count = int(cylinder_count)
    except (TypeError, ValueError):
        return ""

    if count in {3, 4, 5}:
        return f"I{count}"
    if count in {6, 8, 10, 12, 16}:
        return f"V{count}"
    return ""


def is_generic_manual_filename(stem):
    """Return True for placeholder filenames that should not override PDF-based vehicle detection."""
    normalized = compact_vehicle_label(stem)
    return bool(
        re.fullmatch(r"manual\d*", normalized)
        or re.fullmatch(r"ownersmanual\d*", normalized)
        or re.fullmatch(r"operatorsmanual\d*", normalized)
    )


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


def make_reference_patterns(reference=None):
    """Build normalized make/model regex patterns from the optional reference data."""
    if reference is None:
        reference = load_vehicle_reference()

    patterns = []
    for make, models in reference.items():
        make_pattern = r"\b" + re.escape(make.lower()).replace(r"\ ", r"\s+") + r"\b"
        model_patterns = []
        for model in models:
            pattern = r"\b" + re.escape(model.lower()).replace(r"\ ", r"\s+") + r"\b"
            model_patterns.append((model, pattern))
        patterns.append((make, make_pattern, model_patterns))

    return patterns


def find_vehicle_mentions(text, make_filter=None, reference=None):
    """Find known make/model mentions in text with approximate positions."""
    if not text:
        return []

    reference = reference or load_vehicle_reference()
    text_lower = text.lower()
    mentions = []

    for make, make_pattern, model_patterns in make_reference_patterns(reference):
        if make_filter and normalize_vehicle_label(make) != normalize_vehicle_label(make_filter):
            continue

        make_positions = [m.start() for m in re.finditer(make_pattern, text_lower, re.I)]
        if not make_positions:
            continue

        for model, model_pattern in model_patterns:
            for model_match in re.finditer(model_pattern, text_lower, re.I):
                model_pos = model_match.start()
                nearest_make = min(make_positions, key=lambda pos: abs(pos - model_pos))
                mentions.append({
                    "make": make,
                    "model": model,
                    "model_pos": model_pos,
                    "make_pos": nearest_make,
                    "distance": abs(nearest_make - model_pos),
                })

    mentions.sort(key=lambda item: (item["model_pos"], item["distance"], len(item["model"])))
    return mentions


def choose_primary_vehicle_mention(mentions):
    """Pick the strongest make/model mention from early document text."""
    if not mentions:
        return None

    ranked = sorted(
        mentions,
        key=lambda item: (
            item["distance"] > 80,
            item["model_pos"],
            item["distance"],
            -len(item["model"]),
        )
    )
    return ranked[0]


def extract_model_only_targets(text, make=None, reference=None):
    """Find model names explicitly marked as 'Model only' near oil/capacity text."""
    if not text:
        return []

    text_lower = text.lower()
    reference = reference or load_vehicle_reference()
    target_models = []

    oil_anchor_pattern = r"(engine crankcase|oil and filter change|engine oil|oil capacity)"
    oil_anchors = list(re.finditer(oil_anchor_pattern, text_lower, re.I))
    if not oil_anchors:
        return target_models

    candidate_models = {
        mention["model"]
        for mention in find_vehicle_mentions(text, make_filter=make, reference=reference)
    }

    for anchor in oil_anchors:
        context_start = max(0, anchor.start() - 250)
        context_end = min(len(text_lower), anchor.end() + 150)
        context = text_lower[context_start:context_end]
        for model in sorted(candidate_models, key=len, reverse=True):
            model_pattern = r"\b" + re.escape(model.lower()).replace(r"\ ", r"\s+") + r"\s+only\b"
            if re.search(model_pattern, context, re.I) and model not in target_models:
                target_models.append(model)

    return target_models


def build_vehicle_output_targets(filename, raw_text, year, make, model):
    """Return one or more vehicle targets for a PDF, splitting supplements when needed."""
    reference = load_vehicle_reference()
    base_target = [{
        "result_key": filename,
        "year": year,
        "make": make,
        "model": model,
    }]

    explicit_models = extract_model_only_targets(raw_text, make=make, reference=reference)
    if explicit_models and make:
        multi_target = len(explicit_models) > 1
        targets = []
        for explicit_model in explicit_models:
            result_key = filename
            if multi_target:
                stem = filename[:-4] if filename.lower().endswith(".pdf") else filename
                result_key = f"{stem} [{explicit_model}].pdf"
            targets.append({
                "result_key": result_key,
                "year": year,
                "make": make,
                "model": explicit_model,
            })
        return targets

    return base_target


def extract_variant_engine_labels_from_pdf(doc, make=None, model=None):
    """Collect document-native variant labels when manuals expose trims but not engine sizes."""
    if not doc or not model:
        return []

    full_text = ""
    for page in doc:
        full_text += page.get_text() + "\n"

    lines = [re.sub(r"\s+", " ", line.strip()) for line in full_text.split("\n")]
    model_pattern = r"\b" + re.escape(str(model)).replace(r"\ ", r"\s+") + r"\b"
    make_pattern = r"\b" + re.escape(str(make)).replace(r"\ ", r"\s+") + r"\b" if make else ""
    technical_markers = (
        "technical data", "capacities", "weights", "tire inflation",
        "dimensions", "reference", "engine oil", "fuel tank"
    )
    bad_suffixes = {
        "engine", "gasoline engine", "diesel engine", "vehicle", "vehicles",
        "technical data", "reference", "weights", "capacities"
    }

    variants = []
    seen = set()
    for idx, line in enumerate(lines):
        if not line:
            continue
        line_lower = line.lower()
        if not re.search(model_pattern, line, re.I):
            continue

        context_window = " ".join(lines[max(0, idx - 6):min(len(lines), idx + 6)]).lower()
        if not any(marker in context_window for marker in technical_markers):
            continue

        suffix = re.sub(model_pattern, "", line, count=1, flags=re.I).strip(" -:/")
        if make_pattern:
            suffix = re.sub(make_pattern, "", suffix, count=1, flags=re.I).strip(" -:/")
        suffix = re.sub(r"\s+", " ", suffix).strip()
        suffix_lower = suffix.lower()

        if not suffix or suffix_lower in bad_suffixes:
            continue
        if len(suffix.split()) > 4:
            continue
        if not re.search(r"\d", suffix) and not has_engine_signal(suffix_lower):
            continue

        compact_suffix = compact_vehicle_label(suffix)
        if not compact_suffix or compact_suffix in seen:
            continue

        if compact_suffix not in seen:
            seen.add(compact_suffix)
            variants.append(suffix)

    filtered = []
    compact_variants = [(variant, compact_vehicle_label(variant)) for variant in variants]
    for variant, compact_variant in compact_variants:
        if any(
            other_compact != compact_variant
            and compact_variant in other_compact
            and len(other_compact) > len(compact_variant)
            for _, other_compact in compact_variants
        ):
            continue
        filtered.append(variant)

    return filtered


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
    stop_words = set(INVALID_WORDS) | VEHICLE_NAVIGATION_STOP_WORDS | {
        "manual", "owners", "owner", "guide", "service", "maintenance",
        "vehicle", "vehicles", "specification", "specifications", "capacity", "oil",
        "and", "for", "the", "with", "your", "you", "this", "that", "from", "page",
        "pages", "use", "using", "recommended", "information", "summary",
        "gasoline", "engine", "api", "seal", "may", "never", "goes", "below",
        "above", "temperature", "temperatures", "protection", "improved", "economy"
    }

    reference = load_vehicle_reference()
    mentions = find_vehicle_mentions(text, reference=reference)
    primary_mention = choose_primary_vehicle_mention(mentions)
    if primary_mention:
        return year, primary_mention["make"], primary_mention["model"]

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


def extract_engine_specific_oil_map(doc):
    """Map oils to engine sizes from engine-oil specification sections in the manual."""
    engine_oil_map = {}
    full_text = "\n".join(page.get_text("text") for page in doc)
    lines = full_text.splitlines()

    current_engine = None
    for idx, line in enumerate(lines):
        line_lower = line.lower().strip()
        if not line_lower:
            continue

        header_window = " ".join(
            part.strip().lower()
            for part in lines[idx:min(len(lines), idx + 3)]
            if str(part).strip()
        )
        engine_header_match = re.search(
            r"(?:engine oil capacity and\s+specification|capacities and specifications)\s*-\s*(\d{1,2}\.\d)\s*l",
            header_window,
            re.I,
        )
        if engine_header_match:
            current_engine = f"{float(engine_header_match.group(1)):.1f}L"
            engine_oil_map.setdefault(current_engine, [])
            continue

        if current_engine is None:
            continue

        if idx > 0 and "engine oil capacity and" in line_lower:
            continue
        if re.match(r"^(vehicle data|technical data|audio system|accessories|customer assistance|scheduled maintenance)\b", line_lower):
            current_engine = None
            continue
        if "alternative engine oil" in line_lower:
            for look_ahead in range(idx + 1, min(len(lines), idx + 8)):
                alt_line = normalize_ocr_oil_text(lines[look_ahead])
                for base, grade in re.findall(OIL_PATTERN, alt_line, re.I):
                    oil = normalize_oil(f"{base}W-{grade}")
                    if oil not in engine_oil_map[current_engine]:
                        engine_oil_map[current_engine].append(oil)
            continue

        normalized_line = normalize_ocr_oil_text(line)
        for base, grade in re.findall(OIL_PATTERN, normalized_line, re.I):
            oil = normalize_oil(f"{base}W-{grade}")
            if oil not in engine_oil_map[current_engine]:
                engine_oil_map[current_engine].append(oil)

    return engine_oil_map


def extract_model_specific_oil_map(doc, detected_engines=None, model=None):
    """Map explicit model-name oil statements to detected engine variants."""
    full_text = "\n".join(page.get_text("text") for page in doc)
    lines = [re.sub(r"\s+", " ", line.strip()) for line in full_text.splitlines()]
    engine_oil_map = {}

    for idx in range(len(lines)):
        line = lines[idx]
        line_lower = line.lower()
        candidate = None

        if ":" in line and "sae" in line_lower:
            candidate = line
            label_part = candidate.split(":", 1)[0].strip()
            if (not re.search(r"\d", label_part) or candidate.lstrip().startswith("+")) and idx > 0:
                candidate = f"{lines[idx - 1]} {candidate}".strip()
        elif ":" in line and idx + 1 < len(lines) and "sae" in lines[idx + 1].lower():
            candidate = f"{line} {lines[idx + 1]}".strip()
        else:
            continue

        candidate_lower = candidate.lower()
        if "sae" not in candidate_lower or ":" not in candidate:
            continue
        if has_non_engine_oil_context(candidate_lower) or (
            "engine oil" not in candidate_lower and "engine oils" not in candidate_lower
        ):
            continue

        oils = extract_oil_types_from_text(candidate)
        if not oils:
            continue

        label = candidate.split(":", 1)[0].strip()
        if not re.search(r"\d", label):
            continue

        target_engine = match_model_label_to_detected_engine(
            label,
            detected_engines=detected_engines,
            model=model,
        )
        if not target_engine:
            continue

        engine_oil_map.setdefault(target_engine, [])
        for oil in oils:
            if oil not in engine_oil_map[target_engine]:
                engine_oil_map[target_engine].append(oil)

    return engine_oil_map


def extract_engines(text):
    """Extract engine displacements and variants from prose while avoiding capacity rows."""
    engines = []
    text_lower = text.lower()
    
    skip_phrases = [
        "previous generation", "previous model", "prior generation",  
        "towing capacity", "cargo capacity", "payload capacity",
        "weight rating", "gvwr", "gcwr", "optional", "as an option",
        "engine oil", "transaxle fluid", "cooling system", "fuel tank",
        "warranty", "coverage", "owners guide", "owner's guide",
        "diesel engine coverage"
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
        
        sentence_engine_matches = list(re.finditer(ENGINE_PATTERN, sentence_lower))
        for match_idx, m in enumerate(sentence_engine_matches):
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
                    
                    next_engine_start = sentence_engine_matches[match_idx + 1].start() if match_idx + 1 < len(sentence_engine_matches) else len(sentence_lower)
                    window_start = m.start() if match_idx > 0 else max(0, m.start() - 50)
                    window_end = min(next_engine_start, m.end() + 100)
                    window = sentence_lower[window_start:window_end]
                    variant = extract_engine_variant_from_context(window, eng_str)
                    
                    full_eng = eng_str + variant
                    if full_eng not in engines:
                        engines.append(full_eng)
            except (ValueError, TypeError):
                continue
    
    # No static code-to-displacement fallback. Keep only engines found in document text.
    alt_engine_pattern = r"\b(?:[a-z]+\s+)?g?(\d{1,2}\.\d)\s+(?:t-?gdi|tgdi|gdi|mpi|crdi|diesel|turbo)\b"
    for match in re.finditer(alt_engine_pattern, text_lower, re.I):
        try:
            num = float(match.group(1))
        except (ValueError, TypeError):
            continue
        if not is_plausible_engine_displacement(num):
            continue
        eng_str = f"{num:.1f}L"
        context = text_lower[max(0, match.start() - 20):min(len(text_lower), match.end() + 40)]
        variant = extract_engine_variant_from_context(context, eng_str)
        full_eng = eng_str + variant
        if full_eng not in engines:
            engines.append(full_eng)

    for code_engine in extract_engine_code_labels(text):
        if code_engine not in engines:
            engines.append(code_engine)

    return list(dict.fromkeys(engines))


def extract_engines_from_spec_table(text):
    """Extract engines from structured specification tables before falling back to prose."""
    table_engines = []
    lines = text.split('\n')

    for named_engine in extract_named_engine_variants_from_text(text):
        if named_engine not in table_engines:
            table_engines.append(named_engine)
    
    # First find likely engine-spec table headers, then inspect nearby rows.
    spec_header_indices = []
    for idx, line in enumerate(lines):
        line_lower = line.lower().strip()
        
        if is_capacity_or_fluid_row(line_lower):
            continue

        # Keep this strict so "Engine Oil with Filter" is not treated as a spec table.
        has_engine_spec_title = "engine" in line_lower and ("spec" in line_lower or "data" in line_lower)
        table_keyword_pattern = r"\b(?:vin|transaxle|spark|code|gap|cubic\s+inches|firing\s+order|compression\s+ratio)\b"
        lookahead = " ".join(
            lines[next_idx].lower().strip()
            for next_idx in range(idx, min(idx + 6, len(lines)))
        )
        has_engine_table_header = "engine" in line_lower and bool(
            re.search(table_keyword_pattern, lookahead, re.I)
        )

        has_table_keywords = bool(re.search(table_keyword_pattern, line_lower, re.I))
        
        if has_engine_spec_title or has_engine_table_header or (has_table_keywords and len(spec_header_indices) < 10):
            spec_header_indices.append(idx)
    
    for header_idx in spec_header_indices:
        for row_idx in range(header_idx + 1, min(header_idx + 20, len(lines))):
            row = lines[row_idx].strip()
            
            if not row:
                continue
            row_layout_tokens = extract_layout_engine_tokens(row)
            if len(row) < 3 and not row_layout_tokens:
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
                        row_window_start = eng_match.start() if match_idx > 0 else max(0, eng_match.start() - 15)
                        row_window = row_lower[row_window_start:next_engine_start]
                        variant = extract_engine_variant_from_context(row_window, eng_str)
                        
                        full_eng = eng_str + variant
                        if full_eng not in engines_in_row:
                            engines_in_row.append(full_eng)
                except (ValueError, TypeError):
                    continue
            
            if not engines_in_row:
                for alt_match in re.finditer(r"\b(?:[a-z]+\s+)?g?(\d{1,2}\.\d)\s+(?:t-?gdi|tgdi|gdi|mpi|crdi|diesel|turbo)\b", row_lower, re.I):
                    try:
                        eng_size_num = float(alt_match.group(1))
                    except (ValueError, TypeError):
                        continue
                    if not is_plausible_engine_displacement(eng_size_num):
                        continue
                    eng_str = f"{eng_size_num:.1f}L"
                    context = row_lower[max(0, alt_match.start() - 15):min(len(row_lower), alt_match.end() + 40)]
                    variant = extract_engine_variant_from_context(context, eng_str)
                    full_eng = eng_str + variant
                    if full_eng not in engines_in_row:
                        engines_in_row.append(full_eng)

            if not engines_in_row:
                for code_engine in extract_engine_code_labels(row):
                    if code_engine not in engines_in_row:
                        engines_in_row.append(code_engine)

            if not engines_in_row:
                layout_tokens = row_layout_tokens
                row_is_layout_cell = (
                    layout_tokens
                    and re.sub(r"[^a-z0-9]+", "", row_lower) in {
                        re.sub(r"[^a-z0-9]+", "", token.lower())
                        for token in layout_tokens
                    }
                )
                row_has_engine_type_label = "engine type" in row_lower
                header_window = " ".join(lines[header_idx:min(header_idx + 8, len(lines))]).lower()
                row_looks_like_engine_spec_values = bool(
                    layout_tokens
                    and re.search(r"\b(vin|code|transaxle|spark|gap|firing\s+order)\b", header_window, re.I)
                    and (
                        re.search(r"\b(?:automatic|manual)\b", row_lower, re.I)
                        or re.search(r"\b[0-9a-z]\b", row_lower, re.I)
                        or re.search(r"\b\d-\d-\d", row_lower)
                    )
                )
                if row_is_layout_cell or row_has_engine_type_label or row_looks_like_engine_spec_values:
                    for token in layout_tokens:
                        if token not in engines_in_row:
                            engines_in_row.append(token)

            table_engines.extend(engines_in_row)
    
    return list(dict.fromkeys(table_engines))


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
    unmatched = []
    for eng in engines:
        match = re.search(r'(\d+\.?\d*L)', eng)
        if match:
            base = match.group(1)
            if base not in groups:
                groups[base] = []
            groups[base].append(eng)
        else:
            unmatched.append(eng)

    if not groups:
        return engines
    
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

    for eng in unmatched:
        if eng not in result:
            result.append(eng)
    
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
    allowed_displacements = set()
    if all_engines:
        allowed_displacements = {
            round(displacement, 1)
            for displacement in (get_engine_displacement(eng) for eng in all_engines)
            if is_plausible_engine_displacement(displacement)
        }
        for eng in all_engines:
            for token in extract_layout_engine_tokens(eng):
                engine_types_found.add(token)

    # Pass 0: some older manuals list only an engine layout in the
    # specifications table, e.g. "Engine Type" followed by "V6".
    spec_header_indices = []
    for idx, line in enumerate(lines):
        line_lower = line.strip()
        if "engine" in line_lower and ("spec" in line_lower or "vin code" in line_lower or "type" in line_lower):
            spec_header_indices.append(idx)

    for header_idx in spec_header_indices[:12]:
        for row in lines[header_idx:min(header_idx + 20, len(lines))]:
            row_lower = row.strip()
            if not row_lower or is_capacity_or_fluid_row(row_lower):
                continue

            layout_tokens = extract_layout_engine_tokens(row_lower)
            if not layout_tokens:
                continue

            compact_row = re.sub(r"[^a-z0-9]+", "", row_lower)
            layout_compacts = {
                re.sub(r"[^a-z0-9]+", "", token.lower())
                for token in layout_tokens
            }
            if (
                compact_row in layout_compacts
                or "engine type" in row_lower
                or "engine" in " ".join(lines[max(0, header_idx - 2):header_idx + 3])
            ):
                for token in layout_tokens:
                    engine_types_found.add(token)
    
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
                or (allowed_displacements and round(displacement, 1) not in allowed_displacements)
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
            
            mechanical_pattern = r"\b(v\s*-?\s*(?:3|4|5|6|8|10|12|16|20|24)|i\s*-?\s*[3-8]|l\s*-?\s*[3-8]|h\s*-?\s*4|w\s*-?\s*(?:8|12|16)|f\s*-?\s*8|flat[-]?(?:4|6|12)|boxer|boxe|rotary|wankel|ecoboost|turbo|dual[-]?turbo|quad[-]?turbo|sequential[-]?turbo|supercharged|twin[-]?supercharged|twin[-]?turbo|twin[-]?scroll)\b"
            
            for match in re.finditer(mechanical_pattern, descriptor_text):
                match_text = normalize_engine_type_token(match.group())
                engine_types_found.add(match_text)

            for family_token in find_engine_family_tokens(descriptor_text):
                engine_types_found.add(family_token)
            
            inline_pattern = r"\binline\s*[-]?\s*([3-8])\b"
            for match in re.finditer(inline_pattern, descriptor_text):
                digit = match.group(1)
                engine_types_found.add(f"I{digit}")
            
            for match in re.finditer(CYLINDER_PATTERN, descriptor_text):
                token = normalize_cylinder_match(match)
                if token:
                    engine_types_found.add(token)
            for match in re.finditer(r"\b(10|12)\s*[-]?\s*cylinder\b", descriptor_text, re.I):
                engine_types_found.add(f"V{match.group(1)}")

    # Pass 1.5: re-scan the cleaned document around known engine displacements with
    # a wider window so wrapped specs like "3.8L (Code 1) V6" still contribute
    # layout data even when another mention nearby says "Supercharged".
    if all_engines:
        for eng in all_engines:
            displacement = get_engine_displacement(eng)
            if not is_plausible_engine_displacement(displacement):
                continue

            disp_pattern = rf"\b{displacement:.1f}(?:\s*(?:l|liter|litre)\b)?"
            for match in re.finditer(disp_pattern, text_lower, re.I):
                context_start = max(0, match.start() - 25)
                context_end = min(len(text_lower), match.end() + 140)
                context = text_lower[context_start:context_end]
                if not re.search(r"\b(engine|cylinder|displacement|gasoline engines?)\b", context, re.I):
                    continue

                for layout_match in re.finditer(
                    r"\b(v\s*-?\s*(?:3|4|5|6|8|10|12|16|20|24)|i\s*-?\s*[3-8]|l\s*-?\s*[3-8]|h\s*-?\s*4|w\s*-?\s*(?:8|12|16)|f\s*-?\s*8|flat[-]?(?:4|6|12)|boxer|boxe|rotary|wankel)\b",
                    context,
                    re.I
                ):
                    token = normalize_engine_type_token(layout_match.group())
                    if token:
                        engine_types_found.add(token)

                for cylinder_match in re.finditer(CYLINDER_PATTERN, context, re.I):
                    token = normalize_cylinder_match(cylinder_match)
                    if token:
                        engine_types_found.add(token)
                for cylinder_match in re.finditer(r"\b(10|12)\s*[-]?\s*cylinder\b", context, re.I):
                    engine_types_found.add(f"V{cylinder_match.group(1)}")
    
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
        
        if not any(is_layout_engine_type(et) for et in engine_types_found):
            if engine_types_found:
                if max_disp < 1.3:
                    engine_types_found.add("I3")
                elif max_disp < 2.5:
                    engine_types_found.add("I4")
            else:
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

    if any(et in engine_types_found for et in {"V10", "V12", "V16"}):
        engine_types_found -= {"V6", "V8"}
    
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
    
    eng_aliases = engine_code_aliases(engine_str)
    
    for page in doc:
        text = page.get_text()
        text_lower = text.lower()
        
        oil_keywords = r"(?:engine\s+oil|oil\s+with\s+filter|oil\s+capacity|oil\s+change|oil\s+drain|oil\s+fill)"
        
        for match in re.finditer(oil_keywords, text_lower, re.IGNORECASE):
            start = max(0, match.start() - 50)
            end = min(len(text), match.end() + 200)
            window = text[start:end]
            
            window_lower = window.lower()
            if any(alias in window_lower for alias in eng_aliases):
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
        if is_capacity_conversion_engine_match(full_text, m) or is_parenthesized_capacity_conversion(full_text, m):
            continue
        if overlaps_real_capacity_match(full_text, m):
            continue
        eng_size = f"{displacement:.1f}L"
        start = max(0, m.start() - 100)
        end = min(len(full_text), m.end() + 200)
        context = full_text[start:end]
        if any(term in context for term in NON_VEHICLE_ENGINE_CONTEXT):
            continue
        
        engine_matches.append((eng_size, context, m.start()))

    for m in re.finditer(ENGINE_CODE_PATTERN, full_text, re.I):
        start = max(0, m.start() - 100)
        end = min(len(full_text), m.end() + 200)
        context = full_text[start:end]
        full_eng = normalize_engine_code_label(m, context)
        if full_eng:
            engine_matches.append((full_eng.split()[0], context, m.start()))
    
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


def extract_engine_oil_capacity_sections(doc):
    """
    Parse explicit engine-oil capacity sections before broad fallback scans.

    Handles rows such as:
    - Engine oil fill capacity including the oil filter. 1.2 gal (4.6 L)
    - ENGINE OIL CAPACITY AND SPECIFICATION - 2.0L ... All. 5.5 qt (5.2 L)
    - Engine oil (includes filter change) ... All 9.5 quarts (9.0L)
    """
    candidates = []
    full_text = ""
    for page in doc:
        full_text += page.get_text() + "\n"

    lines = full_text.split("\n")

    def clean_line(text):
        """Collapse whitespace in a line while preserving its table text."""
        return re.sub(r"\s+", " ", str(text).strip())

    def capacity_matches_in(text):
        """Return real capacity matches from text, excluding engine-size lookalikes."""
        return [
            m for m in re.finditer(CAPACITY_PATTERN, text, re.I)
            if is_real_capacity_match(text, m)
            and not is_non_capacity_oil_quantity_match(text, m)
        ]

    def engine_key_from_line(text):
        """Extract a plausible engine key from a table line."""
        capacity_matches = capacity_matches_in(text)
        for match in re.finditer(ENGINE_PATTERN, text, re.I):
            if (
                is_capacity_conversion_engine_match(text, match)
                or is_parenthesized_capacity_conversion(text, match)
                or overlaps_real_capacity_match(text, match, capacity_matches)
            ):
                continue
            try:
                displacement = float(match.group(1))
            except (TypeError, ValueError):
                continue
            if is_plausible_engine_displacement(displacement):
                return f"{displacement:.1f}L"
        return None

    def engine_key_from_capacity_heading(text):
        """Extract an engine key from a capacities/specifications heading."""
        heading_match = re.search(
            r"capacities\s+and\s+specifications\s*-\s*(\d{1,2}\.\d)\s*l[^\n]*",
            text,
            re.I,
        )
        if not heading_match:
            return None

        try:
            displacement = float(heading_match.group(1))
        except (TypeError, ValueError):
            return None
        if not is_plausible_engine_displacement(displacement):
            return None

        base_engine = f"{displacement:.1f}L"
        variant = extract_engine_variant_from_context(heading_match.group(0), base_engine)
        return base_engine + variant

    def capacity_from_text(text):
        """Build one validated capacity record from a text fragment."""
        matches = capacity_matches_in(text)
        selected = choose_preferred_capacity_match(matches)
        if not selected:
            return None
        try:
            capacity = build_capacity_record_from_matches(selected, matches)
        except (ValueError, TypeError, AttributeError):
            return None
        q = capacity.get("quarts")
        if ENGINE_OIL_CAPACITY_MIN_QT <= q <= ENGINE_OIL_CAPACITY_MAX_QT:
            return capacity
        return None

    def capacity_from_matches(matches):
        """Choose the best in-range capacity from an existing match list."""
        valid_matches = []
        for match in matches:
            try:
                q, _ = to_quarts_liters(match.group(1), match.group(2))
            except (TypeError, ValueError, AttributeError):
                continue
            if ENGINE_OIL_CAPACITY_MIN_QT <= q <= ENGINE_OIL_CAPACITY_MAX_QT:
                valid_matches.append(match)

        selected = choose_preferred_capacity_match(valid_matches)
        if not selected:
            return None
        try:
            return build_capacity_record_from_matches(selected, valid_matches)
        except (TypeError, ValueError, AttributeError):
            return None

    def nearby_previous_capacity_for_item(label_idx):
        """Find a capacity immediately above an engine-oil row when the table is inverted."""
        prior_lines = []
        earliest_prior_idx = label_idx
        for prior_idx in range(label_idx - 1, max(-1, label_idx - 4), -1):
            prior_clean = clean_line(lines[prior_idx])
            prior_lower = prior_clean.lower()
            if not prior_lower:
                continue
            if not capacity_matches_in(prior_lower):
                break
            prior_lines.insert(0, prior_clean)
            earliest_prior_idx = prior_idx

        if not prior_lines:
            return None, None

        # In application/capacity tables, the values immediately before an
        # engine-oil row can belong to the previous fluid row. Reject those
        # before they get shared as a generic engine-oil capacity.
        prior_label = ""
        for label_lookup_idx in range(earliest_prior_idx - 1, max(-1, earliest_prior_idx - 5), -1):
            label_clean = clean_line(lines[label_lookup_idx])
            label_lower = label_clean.lower()
            if not label_lower or capacity_matches_in(label_lower):
                continue
            prior_label = label_lower
            break

        generic_prior_labels = {
            "application", "applications", "capacity", "capacities",
            "metric", "english", "metric english", "quantity", "item",
            "items", "vehicle data", "technical data",
        }
        if prior_label:
            prior_label_compact = re.sub(r"[^a-z]+", " ", prior_label).strip()
            if any(term in prior_label for term in ENGINE_OIL_STOP_TERMS):
                return None, None
            if (
                "engine oil" not in prior_label
                and "motor oil" not in prior_label
                and "oil and filter" not in prior_label
                and prior_label_compact not in generic_prior_labels
            ):
                return None, None

        prior_text = " ".join(prior_lines)
        return capacity_from_matches(capacity_matches_in(prior_text)), prior_text

    def forward_capacity_for_oil_label(label_idx):
        """Scan following rows for the first engine-oil capacity after a label."""
        window_lines = []
        for follow_idx in range(label_idx + 1, min(len(lines), label_idx + 12)):
            follow_clean = clean_line(lines[follow_idx])
            follow_lower = follow_clean.lower()
            if not follow_lower:
                continue

            if any(term in follow_lower for term in [
                "transmission", "transaxle", "differential", "power steering",
                "brake fluid", "washer", "refrigerant", "wheel nut", "transfer case"
            ]):
                break

            window_lines.append(follow_clean)
            capacity = capacity_from_matches(capacity_matches_in(" ".join(window_lines)))
            if capacity:
                return capacity, " ".join(window_lines)

        return None, None

    def find_capacity_after_label(start_idx, target_field):
        """Find the capacity tied to an oil label near the current line."""
        current_line = clean_line(lines[start_idx])
        current_lower = current_line.lower()

        def nearest_prior_label(index):
            """Return the nearest previous non-capacity label before an index."""
            for prior_idx in range(index, -1, -1):
                prior_clean = clean_line(lines[prior_idx])
                prior_lower = prior_clean.lower()
                if not prior_lower:
                    continue
                if capacity_from_text(prior_lower):
                    continue
                return prior_lower
            return ""

        if detect_capacity_field(current_lower):
            backward_matches = []
            for follow_idx in range(max(0, start_idx - 2), start_idx):
                candidate_line = clean_line(lines[follow_idx])
                candidate_lower = candidate_line.lower()
                if not candidate_lower:
                    continue
                capacity = capacity_from_text(candidate_lower)
                if capacity:
                    backward_matches.append((capacity, candidate_line, follow_idx))

            if backward_matches:
                prior_label = nearest_prior_label(backward_matches[0][2] - 1)
                generic_headers = {"capacity", "capacities", "item", "items", "quantity", "metric", "english", "variant"}
                if not prior_label or prior_label in generic_headers:
                    return backward_matches[-1][0], backward_matches[-1][1]

        for follow_idx in range(start_idx, min(len(lines), start_idx + 3)):
            candidate_line = clean_line(lines[follow_idx])
            candidate_lower = candidate_line.lower()
            if not candidate_lower:
                continue
            if follow_idx > start_idx and detect_capacity_field(candidate_lower):
                break
            if any(term in candidate_lower for term in ENGINE_OIL_STOP_TERMS + ["adding engine oil", "dipstick"]):
                break
            capacity = capacity_from_text(candidate_lower)
            if capacity:
                return capacity, candidate_line
        return None, None

    def is_page_footer_or_noise(text):
        """Detect footer/header lines that should not be parsed as data."""
        text_lower = text.lower()
        return bool(
            re.fullmatch(r"\d{1,4}", text_lower)
            or "edition date" in text_lower
            or "first-printing" in text_lower
            or "second printing" in text_lower
            or "capacities and specifications" == text_lower
        )

    def heading_context_at(index, before=3, after=3):
        """Return nearby heading context around a line index."""
        start = max(0, index - before)
        end = min(len(lines), index + after + 1)
        return " ".join(clean_line(line).lower() for line in lines[start:end])

    current_engine = None

    # Row-oriented table pattern:
    # Engine Oil with Filter
    # 1.4L L4
    # 4.0 L
    # 4.2 qt
    table_engine = None
    for line_idx, line in enumerate(lines):
        line_clean = clean_line(line)
        line_lower = line_clean.lower()
        if not line_lower:
            continue

        heading_engine = engine_key_from_capacity_heading(line_clean)
        if heading_engine:
            table_engine = heading_engine

        if "engine oil" not in line_lower or "life" in line_lower or "pressure" in line_lower:
            continue
        if "dipstick" in line_lower or "adding engine oil" in line_lower:
            continue
        if (
            re.search(OIL_PATTERN, line_lower, re.I)
            and "capacity" not in line_lower
            and not any(label in line_lower for label in WITH_FILTER_LABELS + WITHOUT_FILTER_LABELS)
        ):
            continue
        if any(term in line_lower for term in ["oil filter part", "engine oil filler cap"]):
            continue

        context = heading_context_at(line_idx, before=8, after=2)
        in_capacity_table = (
            table_engine is not None
            or ("capacities" in context and "specifications" in context)
            or ("capacity item" in re.sub(r"[^a-z]+", " ", context))
        )
        if not in_capacity_table:
            continue

        capacity, capacity_text = nearby_previous_capacity_for_item(line_idx)
        if not capacity:
            capacity, capacity_text = forward_capacity_for_oil_label(line_idx)
        if not capacity:
            continue

        field = detect_capacity_field(line_lower) or "with_filter"
        target_engine = table_engine or "unknown_engine"
        evidence_text = f"{line_clean} {capacity_text or ''}".strip()
        candidates.append({
            "engine": target_engine,
            "field": field,
            "capacity": capacity,
            "score": score_capacity_candidate(
                "engine oil capacity " + evidence_text,
                target_field=field,
                engine_key=target_engine,
            ) + 12,
        })

    row_label_patterns = [
        "engine oil with filter",
        "engine oil fill capacity including the oil filter",
        "oil and filter change",
    ]
    for line_idx, line in enumerate(lines):
        line_clean = clean_line(line)
        line_lower = line_clean.lower()
        if not any(label in line_lower for label in row_label_patterns):
            continue

        direct_capacity, direct_text = nearby_previous_capacity_for_item(line_idx)
        if not direct_capacity:
            direct_capacity, direct_text = forward_capacity_for_oil_label(line_idx)
        if direct_capacity:
            candidates.append({
                "engine": "unknown_engine",
                "field": "with_filter",
                "capacity": direct_capacity,
                "score": score_capacity_candidate(
                    "engine oil with filter " + (direct_text or line_clean),
                    target_field="with_filter",
                    engine_key="unknown_engine",
                ) + 10,
            })

        vertical_engines = []
        vertical_caps = []
        for follow_idx in range(line_idx + 1, min(len(lines), line_idx + 12)):
            follow_clean = clean_line(lines[follow_idx])
            follow_lower = follow_clean.lower()
            if not follow_lower:
                continue
            if any(term in follow_lower for term in ENGINE_OIL_STOP_TERMS):
                break
            if re.match(r"^(fuel tank|wheel nut|transfer case|cooling system|air conditioning|technical data|vehicle data)\b", follow_lower):
                break

            for code_engine in extract_engine_code_labels(follow_clean):
                if code_engine not in [item[0] for item in vertical_engines]:
                    vertical_engines.append((code_engine, follow_idx))

            base_engine = engine_key_from_line(follow_lower)
            if base_engine:
                variant = extract_engine_variant_from_context(follow_lower, base_engine)
                engine_key = base_engine + variant
                if engine_key not in [item[0] for item in vertical_engines]:
                    vertical_engines.append((engine_key, follow_idx))

            cap = capacity_from_text(follow_lower)
            if cap:
                vertical_caps.append((cap, follow_idx, follow_clean))

        if (
            len(vertical_engines) >= 2
            and len(vertical_caps) >= len(vertical_engines)
            and vertical_engines[-1][1] < vertical_caps[0][1]
        ):
            for (engine_key, _), (capacity, _, cap_text) in zip(vertical_engines, vertical_caps):
                candidates.append({
                    "engine": engine_key,
                    "field": "with_filter",
                    "capacity": capacity,
                    "score": score_capacity_candidate("engine oil with filter " + cap_text, target_field="with_filter", engine_key=engine_key) + 10,
                })

        pending_engine = None
        pending_matches = []
        for follow_idx in range(line_idx + 1, min(len(lines), line_idx + 12)):
            follow_clean = clean_line(lines[follow_idx])
            follow_lower = follow_clean.lower()
            if not follow_lower:
                continue
            if any(term in follow_lower for term in ENGINE_OIL_STOP_TERMS):
                break
            if re.match(r"^(fuel tank|wheel nut|transfer case|cooling system|air conditioning|technical data|vehicle data)\b", follow_lower):
                break

            engine_row_match = re.search(r"\b(\d{1,2}\.\d)\s*(?:l|liter|litre)\b\s+([a-z0-9-]+)", follow_lower, re.I)
            engine_key = None
            if engine_row_match:
                token_after = engine_row_match.group(2)
                if re.fullmatch(r"(?:l[3-8]|i[3-8]|v\d|ecoboost|gdi|turbo|supercharged|dohc|sohc)", token_after, re.I):
                    displacement = float(engine_row_match.group(1))
                    if is_plausible_engine_displacement(displacement):
                        base_engine = f"{displacement:.1f}L"
                        variant = extract_engine_variant_from_context(follow_lower, base_engine)
                        engine_key = base_engine + variant
                elif re.fullmatch(r"engine|engines", token_after, re.I):
                    displacement = float(engine_row_match.group(1))
                    if is_plausible_engine_displacement(displacement):
                        base_engine = f"{displacement:.1f}L"
                        variant = extract_engine_variant_from_context(follow_lower, base_engine)
                        engine_key = base_engine + variant

            if not engine_key and "engine" in follow_lower:
                base_engine = engine_key_from_line(follow_lower)
                if base_engine:
                    variant = extract_engine_variant_from_context(follow_lower, base_engine)
                    engine_key = base_engine + variant

            if not engine_key:
                code_labels = extract_engine_code_labels(follow_clean)
                if code_labels:
                    engine_key = code_labels[0]

            if engine_key and has_engine_signal(follow_lower) and not is_capacity_or_fluid_row(follow_lower):
                if pending_engine and pending_matches:
                    selected = choose_preferred_capacity_match(pending_matches)
                    if selected:
                        capacity = build_capacity_record_from_matches(selected, pending_matches)
                        candidates.append({
                            "engine": pending_engine,
                            "field": "with_filter",
                            "capacity": capacity,
                            "score": score_capacity_candidate("engine oil with filter " + " ".join(m.group(0) for m in pending_matches), target_field="with_filter", engine_key=pending_engine) + 8,
                        })
                pending_engine = engine_key
                pending_matches = []
                continue

            if pending_engine:
                for match in capacity_matches_in(follow_lower):
                    pending_matches.append(match)

        if pending_engine and pending_matches:
            selected = choose_preferred_capacity_match(pending_matches)
            if selected:
                capacity = build_capacity_record_from_matches(selected, pending_matches)
                candidates.append({
                    "engine": pending_engine,
                    "field": "with_filter",
                    "capacity": capacity,
                    "score": score_capacity_candidate("engine oil with filter " + " ".join(m.group(0) for m in pending_matches), target_field="with_filter", engine_key=pending_engine) + 8,
                })

    for line_idx, line in enumerate(lines):
        line_clean = clean_line(line)
        line_lower = line_clean.lower()
        if not line_lower:
            continue

        context = heading_context_at(line_idx, before=2, after=1)
        line_engine = engine_key_from_line(line_lower)
        if line_engine and not any(term in context for term in ENGINE_OIL_STOP_TERMS):
            if (
                "engine oil capacity" in context
                or ("capacities" in context and "specifications" in context)
            ):
                current_engine = line_engine

        if any(term in line_lower for term in ENGINE_OIL_STOP_TERMS):
            continue
        if "adding engine oil" in line_lower or "dipstick" in line_lower:
            continue
        if is_page_footer_or_noise(line_lower):
            continue

        label_context = " ".join(
            clean_line(x).lower()
            for x in lines[line_idx:min(len(lines), line_idx + 4)]
            if clean_line(x)
        )

        target_field = None
        if any(phrase in label_context for phrase in [
            "including the oil filter",
            "includes filter change",
            "includes filter",
            "including filter",
            "with oil filter",
            "with filter",
            "variant including"
        ]):
            target_field = "with_filter"
        elif "excluding the oil filter" in label_context or "without filter" in label_context:
            target_field = "without_filter"

        if not target_field:
            continue

        current_line_has_filter_label = any(phrase in line_lower for phrase in [
            "including the oil filter",
            "includes filter change",
            "includes filter",
            "including filter",
            "with oil filter",
            "with filter",
            "variant including",
            "excluding the oil filter",
            "without filter",
        ])
        if not current_line_has_filter_label and "engine oil" not in line_lower and "motor oil" not in line_lower:
            continue

        if capacity_from_text(line_lower) and target_field and not any(phrase in line_lower for phrase in [
            "including the oil filter",
            "includes filter change",
            "includes filter",
            "including filter",
            "with oil filter",
            "with filter",
            "variant including",
            "excluding the oil filter",
            "without filter",
        ]):
            continue

        capacity = None
        capacity_text = None
        if current_line_has_filter_label:
            capacity, capacity_text = find_capacity_after_label(line_idx, target_field)
        if not capacity:
            row_text_options = [line_clean]
            row_text_options.extend([
                " ".join(clean_line(x) for x in lines[line_idx:min(len(lines), line_idx + 4)] if clean_line(x)),
                " ".join(clean_line(x) for x in lines[line_idx:min(len(lines), line_idx + 10)] if clean_line(x)),
            ])
            for row_text in row_text_options:
                row_lower = row_text.lower()
                if any(term in row_lower for term in ENGINE_OIL_STOP_TERMS + ["adding engine oil", "dipstick"]):
                    continue
                capacity = capacity_from_text(row_lower)
                if capacity:
                    capacity_text = row_text
                    break

        if not capacity:
            continue

        target_engine = current_engine or "unknown_engine"
        evidence_text = " ".join(part for part in [line_clean, capacity_text] if part)
        candidates.append({
            "engine": target_engine,
            "field": target_field,
            "capacity": capacity,
            "score": score_capacity_candidate(evidence_text, target_field=target_field, engine_key=target_engine),
        })

    selected_caps = select_best_capacity_candidates(candidates)
    explicit_engine_keys = [key for key in selected_caps.keys() if key != "unknown_engine"]
    if explicit_engine_keys and "unknown_engine" in selected_caps:
        selected_caps.pop("unknown_engine", None)
    return selected_caps


def extract_engine_capacities(doc):
    """Extract engine-oil capacities from explicit oil sections and ignore other fluid tables."""
    engine_caps = extract_engine_oil_capacity_sections(doc)
    full_text = ""
    for page in doc:
        full_text += page.get_text() + "\n"
    targeted_spec_ocr_text = extract_targeted_spec_ocr_text(doc)
    if targeted_spec_ocr_text.strip():
        full_text += targeted_spec_ocr_text + "\n"
    lines = full_text.split('\n')

    if engine_caps:
        if has_ambiguous_model_oil_capacity_table(lines) and set(engine_caps.keys()) == {"unknown_engine"}:
            return {}
        return engine_caps

    refill_application_caps = extract_refill_application_engine_oil_capacities(full_text)
    if refill_application_caps:
        return refill_application_caps

    image_table_caps = extract_image_table_engine_oil_capacity(doc)
    if image_table_caps:
        return image_table_caps

    candidates = []

    ordered_ocr_capacity = extract_ordered_engine_oil_capacity_from_text(full_text)
    if ordered_ocr_capacity:
        candidates.append({
            "engine": "unknown_engine",
            "field": "with_filter",
            "capacity": ordered_ocr_capacity,
            "score": score_capacity_candidate(
                "engine oil with filter " + str(ordered_ocr_capacity.get("quarts")) + " quarts",
                target_field="with_filter",
                engine_key="unknown_engine",
            ) + 8,
        })
    
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

                selected_cap = choose_preferred_capacity_match(capacity_matches)

                try:
                    q, l = to_quarts_liters(selected_cap.group(1), selected_cap.group(2))
                    if ENGINE_OIL_CAPACITY_MIN_QT <= q <= ENGINE_OIL_CAPACITY_MAX_QT:
                        evidence_lines = [
                            x.strip() for x in oil_section_lines[section_line_idx:section_line_idx + 3]
                            if x.strip()
                        ]
                        evidence_text = " ".join(evidence_lines) or section_line
                        field = detect_capacity_field(evidence_text) or "with_filter"
                        candidates.append({
                            "engine": "unknown_engine",
                            "field": field,
                            "capacity": {"quarts": q, "liters": l},
                            "score": score_capacity_candidate(evidence_text, target_field=field, engine_key="unknown_engine"),
                        })
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
                
                capacity_matches_after_engine = [
                    cm for cm in capacity_matches
                    if cm.start() >= eng_match.end()
                ]
                selected_cap = choose_preferred_capacity_match(capacity_matches_after_engine)
                
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

                        selected_cap = choose_preferred_capacity_match(next_capacity_matches)
                        break

                if not selected_cap:
                    continue
                
                q, l = to_quarts_liters(selected_cap.group(1), selected_cap.group(2))
                
                if ENGINE_OIL_CAPACITY_MIN_QT <= q <= ENGINE_OIL_CAPACITY_MAX_QT:
                    evidence_lines = [
                        x.strip() for x in oil_section_lines[max(0, section_line_idx - 1):section_line_idx + 3]
                        if x.strip()
                    ]
                    evidence_text = " ".join(evidence_lines) or section_line
                    field = detect_capacity_field(evidence_text) or "with_filter"
                    candidates.append({
                        "engine": eng_size,
                        "field": field,
                        "capacity": {"quarts": q, "liters": l},
                        "score": score_capacity_candidate(evidence_text, target_field=field, engine_key=eng_size),
                    })
            except (ValueError, TypeError, AttributeError):
                continue
    
    selected = select_best_capacity_candidates(candidates)
    if has_ambiguous_model_oil_capacity_table(lines) and set(selected.keys()) == {"unknown_engine"}:
        return {}
    return selected


def extract_columnar_model_capacity_table(doc):
    """Parse technical-data tables that map model columns to engine-oil capacities."""
    full_text = ""
    for page in doc:
        full_text += page.get_text() + "\n"

    lines = [re.sub(r"\s+", " ", line.strip()) for line in full_text.split("\n")]
    model_to_engine = {}
    model_to_layout = {}

    for idx, line in enumerate(lines):
        if "displacement" not in line.lower():
            continue

        model_tokens = []
        for probe in lines[max(0, idx - 4):idx]:
            probe_lower = probe.lower()
            if not probe or probe_lower.startswith("repairs") or probe_lower.startswith("index"):
                continue
            if any(ch.isdigit() for ch in probe) and not probe_lower.startswith(("194n", "195n", "196n", "197n")):
                model_tokens.append(probe)
        if len(model_tokens) < 2:
            continue

        displacement_values = []
        for probe in lines[idx:min(len(lines), idx + 10)]:
            displacement_values.extend(re.findall(r"\((\d{3,5})\)", probe))
            if len(displacement_values) >= len(model_tokens):
                break

        if len(displacement_values) < len(model_tokens):
            continue

        cylinder_values = []
        for probe_idx in range(idx, min(len(lines), idx + 12)):
            if "number of cylinders" not in lines[probe_idx].lower():
                continue
            for probe in lines[probe_idx + 1:min(len(lines), probe_idx + 10)]:
                probe_clean = probe.strip()
                if re.fullmatch(r"\d{1,2}", probe_clean):
                    cylinder_values.append(probe_clean)
                if len(cylinder_values) >= len(model_tokens):
                    break
            break

        for model_token, cc_value in zip(model_tokens, displacement_values):
            try:
                liters = round(int(cc_value) / 1000.0, 1)
            except (TypeError, ValueError):
                continue
            if is_plausible_engine_displacement(liters):
                model_key = compact_vehicle_label(model_token)
                model_to_engine[model_key] = f"{liters:.1f}L"

        if len(cylinder_values) >= len(model_tokens):
            for model_token, cylinder_count in zip(model_tokens, cylinder_values):
                layout_token = cylinder_count_to_layout_token(cylinder_count)
                if layout_token:
                    model_to_layout[compact_vehicle_label(model_token)] = layout_token

    if not model_to_engine:
        return {}

    candidates = []
    sorted_models = sorted(model_to_engine.items(), key=lambda item: len(item[0]), reverse=True)

    for idx, line in enumerate(lines):
        line_lower = line.lower()
        if "engine oil" not in line_lower or "filter" not in line_lower:
            continue

        block_lines = [probe for probe in lines[idx:min(len(lines), idx + 8)] if probe]
        for line_index, probe in enumerate(block_lines):
            capacity_matches = [
                match for match in re.finditer(CAPACITY_PATTERN, probe, re.I)
                if is_real_capacity_match(probe, match)
            ]
            target_engine = None
            local_context = " ".join(block_lines[line_index:min(len(block_lines), line_index + 3)])
            context_key = compact_vehicle_label(local_context)
            for model_key, engine_key in sorted_models:
                if model_key and model_key in context_key:
                    target_engine = engine_key
                    break
            if not target_engine:
                continue

            matched_model_key = next(
                (model_key for model_key, _ in sorted_models if model_key and model_key in context_key),
                ""
            )
            layout_token = model_to_layout.get(matched_model_key, "")
            if layout_token and layout_token not in target_engine.split():
                target_engine = f"{target_engine} {layout_token}"

            capacity = None
            if capacity_matches:
                selected = choose_preferred_capacity_match(capacity_matches)
                if not selected:
                    continue
                try:
                    capacity = build_capacity_record_from_matches(selected, capacity_matches)
                except (ValueError, TypeError, AttributeError):
                    continue
            else:
                units_context = " ".join(block_lines[max(0, line_index - 4):line_index + 1]).lower()
                paired_values = re.search(r"(\d+\.?\d*)\s*\((\d+\.?\d*)\)", probe)
                if paired_values and "quart" in units_context and "liter" in units_context:
                    try:
                        capacity = {
                            "quarts": float(paired_values.group(1)),
                            "liters": float(paired_values.group(2)),
                        }
                    except (ValueError, TypeError):
                        capacity = None
            if not capacity:
                continue

            q = capacity.get("quarts")
            if not (ENGINE_OIL_CAPACITY_MIN_QT <= q <= ENGINE_OIL_CAPACITY_MAX_QT):
                continue

            candidates.append({
                "engine": target_engine,
                "field": "with_filter",
                "capacity": capacity,
                "score": score_capacity_candidate(local_context, target_field="with_filter", engine_key=target_engine) + 10,
            })

    return select_best_capacity_candidates(candidates)


def extract_model_named_capacity_tables(doc, detected_engines=None, model=None):
    """Parse row-based oil-capacity tables that list one or more model labels before each quantity."""
    full_text = ""
    for page in doc:
        full_text += page.get_text() + "\n"

    lines = [re.sub(r"\s+", " ", line.strip()) for line in full_text.split("\n")]
    candidates = []
    table_headers = ("engine oil filling quantity",)
    stop_markers = (
        "notes on brake fluid",
        "coolant",
        "engine coolant",
        "notes on coolant",
        "windshield washer",
        "brake fluid",
        "refrigerant",
        "fuel tank",
        "tank content",
    )

    def is_model_row(text):
        """Return True when a table row names a target model/variant, not a capacity."""
        text_lower = text.lower()
        if not text_lower or text_lower in {"model", "quantity", "filling quantity"}:
            return False
        if any(marker in text_lower for marker in stop_markers):
            return False
        if text_lower.startswith(("the following values", "* ", "to achieve ", "only use ", "recommended ", "not for ", "plug-in hybrid")):
            return False
        if re.fullmatch(r"229\.\d+\*?(?:,\s*229\.\d+\*?)*", text_lower):
            return False
        if re.search(CAPACITY_PATTERN, text_lower, re.I):
            return False
        if len(text.split()) > 12:
            return False
        compact = compact_vehicle_label(canonicalize_engine_variant_label(text, model=model))
        if detected_engines:
            for engine in detected_engines:
                engine_compact = compact_vehicle_label(canonicalize_engine_variant_label(engine, model=model))
                engine_bodyless = compact_vehicle_label(strip_parenthetical_body_style(canonicalize_engine_variant_label(engine, model=model)))
                if compact and compact in {engine_compact, engine_bodyless}:
                    return True
        return bool(re.search(r"\d", text) and any(token in text_lower for token in ["4matic", "amg", "hybrid", "eq", "awd", "xdrive", "quattro", "fwd", "rwd"]))

    for idx, line in enumerate(lines):
        line_lower = line.lower()
        if not any(header in line_lower for header in table_headers):
            continue

        pending_models = []
        for follow_idx in range(idx + 1, min(len(lines), idx + 20)):
            follow = lines[follow_idx]
            follow_lower = follow.lower()
            if not follow_lower:
                continue
            if follow_idx > idx + 1 and any(header in follow_lower for header in table_headers):
                break
            if any(marker in follow_lower for marker in stop_markers):
                break
            if follow_lower in {"model", "quantity", "filling quantity"}:
                continue
            if follow_lower.startswith(("the following values refer to", "engine oil quality and filling quantity")):
                continue

            capacity_matches = [
                match for match in re.finditer(CAPACITY_PATTERN, follow, re.I)
                if is_real_capacity_match(follow_lower, match)
            ]
            if capacity_matches and pending_models:
                selected = choose_preferred_capacity_match(capacity_matches)
                if not selected:
                    pending_models = []
                    continue
                try:
                    capacity = build_capacity_record_from_matches(selected, capacity_matches)
                except (ValueError, TypeError, AttributeError):
                    pending_models = []
                    continue

                q = capacity.get("quarts")
                if isinstance(q, (int, float)) and ENGINE_OIL_CAPACITY_MIN_QT <= q <= ENGINE_OIL_CAPACITY_MAX_QT:
                    for pending_model in pending_models:
                        target_engine = match_model_label_to_detected_engine(
                            pending_model,
                            detected_engines=detected_engines,
                            model=model,
                        )
                        evidence_text = f"{pending_model} {follow}".strip()
                        candidates.append({
                            "engine": target_engine,
                            "field": "with_filter",
                            "capacity": capacity,
                            "score": score_capacity_candidate(
                                f"engine oil filling quantity {evidence_text}",
                                target_field="with_filter",
                                engine_key=target_engine,
                            ) + 12,
                        })
                pending_models = []
                continue

            if is_model_row(follow):
                pending_models.append(follow)
            elif pending_models and any(marker in follow_lower for marker in ["plug-in hybrid", "not for plug-in hybrid"]):
                pending_models = []

    return select_best_capacity_candidates(candidates)


def extract_recommended_lubricants_capacity(doc):
    """Parse generic recommended-lubricants tables that list one engine-oil row with volume."""
    full_text = ""
    for page in doc:
        full_text += page.get_text() + "\n"

    lines = [re.sub(r"\s+", " ", line.strip()) for line in full_text.split("\n")]
    stop_labels = {
        "automatic transmission fluid",
        "coolant",
        "brake fluid",
        "rear differential oil",
        "transfer case oil",
        "fuel",
        "washer fluid",
        "windshield washer fluid",
    }

    for idx, line in enumerate(lines):
        line_lower = line.lower()
        if "recommended lubricants and capacities" not in line_lower:
            continue

        block = []
        for follow_line in lines[idx:min(len(lines), idx + 35)]:
            follow_lower = follow_line.lower()
            if not follow_lower:
                continue
            if block and follow_lower in stop_labels:
                break
            block.append(follow_line)

        block_text = " ".join(block)
        block_lower = block_text.lower()
        if "engine oil" not in block_lower:
            continue

        capacity_matches = [
            m for m in re.finditer(CAPACITY_PATTERN, block_text, re.I)
            if is_real_capacity_match(block_text, m)
        ]
        selected = choose_preferred_capacity_match(capacity_matches)
        if not selected:
            continue

        try:
            capacity = build_capacity_record_from_matches(selected, capacity_matches)
        except (ValueError, TypeError, AttributeError):
            continue

        q = capacity.get("quarts")
        if not (ENGINE_OIL_CAPACITY_MIN_QT <= q <= ENGINE_OIL_CAPACITY_MAX_QT):
            continue

        return {
            "unknown_engine": {
                "with_filter": capacity,
                "without_filter": None,
            }
        }

    return {}


def extract_explicit_shared_engine_oil_capacity(doc):
    """Find shared rows like "Engine Oil with Filter 8.0 L 8.5 qt" before broad fallbacks."""
    full_text = ""
    for page in doc:
        full_text += page.get_text() + "\n"

    lines = full_text.split("\n")
    candidates = []
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

        selected = choose_preferred_capacity_match(capacity_matches)

        try:
            q, l = to_quarts_liters(selected.group(1), selected.group(2))
        except (ValueError, TypeError, AttributeError):
            continue

        if ENGINE_OIL_CAPACITY_MIN_QT <= q <= ENGINE_OIL_CAPACITY_MAX_QT:
            field = detect_capacity_field(window) or "with_filter"
            candidates.append({
                "engine": "unknown_engine",
                "field": field,
                "capacity": {"quarts": q, "liters": l},
                "score": score_capacity_candidate(window, target_field=field, engine_key="unknown_engine"),
            })

    selected = select_best_capacity_candidates(candidates)
    return selected.get("unknown_engine")


def extract_fallback_capacity(doc):
    """Find a generic oil capacity when no engine-specific capacity can be extracted."""
    full_text = ""
    for page in doc:
        full_text += page.get_text() + "\n"
    targeted_spec_ocr_text = extract_targeted_spec_ocr_text(doc)
    if targeted_spec_ocr_text.strip():
        full_text += targeted_spec_ocr_text + "\n"
    lines = full_text.split("\n")

    section_caps = extract_engine_oil_capacity_sections(doc)
    if section_caps:
        if has_ambiguous_model_oil_capacity_table(lines) and set(section_caps.keys()) == {"unknown_engine"}:
            return None
        if "unknown_engine" in section_caps:
            return section_caps["unknown_engine"]
        if len(section_caps) == 1:
            return next(iter(section_caps.values()))
        return None

    if has_ambiguous_model_oil_capacity_table(lines):
        return None

    explicit_shared = extract_explicit_shared_engine_oil_capacity(doc)
    if explicit_shared:
        return explicit_shared

    image_table_caps = extract_image_table_engine_oil_capacity(doc)
    if image_table_caps.get("unknown_engine"):
        return image_table_caps["unknown_engine"]

    recommended_lubricants_caps = extract_recommended_lubricants_capacity(doc)
    if recommended_lubricants_caps.get("unknown_engine"):
        return recommended_lubricants_caps["unknown_engine"]

    ordered_ocr_capacity = extract_ordered_engine_oil_capacity_from_text(full_text)
    if ordered_ocr_capacity:
        return {"with_filter": ordered_ocr_capacity, "without_filter": None}

    candidates = []

    text_sources = [page.get_text() for page in doc]
    if targeted_spec_ocr_text.strip():
        text_sources.append(targeted_spec_ocr_text)

    for source_text in text_sources:
        text = source_text.lower()
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
            selected = choose_preferred_capacity_match(capacity_matches) or wf_m

            q, l = to_quarts_liters(selected.group(1), selected.group(2))
            if ENGINE_OIL_CAPACITY_MIN_QT <= q <= ENGINE_OIL_CAPACITY_MAX_QT:
                candidates.append({
                    "engine": "unknown_engine",
                    "field": "with_filter",
                    "capacity": {"quarts": q, "liters": l},
                    "score": score_capacity_candidate(window, target_field="with_filter", engine_key="unknown_engine"),
                })
        
        lines = text.splitlines()
        for idx, line in enumerate(lines):
            line_lower = line.lower()
            if (
                "engine oil with filter" in line_lower
                or "engine oil fill capacity including the oil filter" in line_lower
                or "oil and filter change" in line_lower
            ):
                for look_idx in range(max(0, idx - 2), min(len(lines), idx + 4)):
                    look_line = lines[look_idx].strip()
                    if not look_line:
                        continue
                    for m in re.finditer(CAPACITY_PATTERN, look_line, re.I):
                        if not is_real_capacity_match(look_line, m):
                            continue
                        q, l = to_quarts_liters(m.group(1), m.group(2))
                        if ENGINE_OIL_CAPACITY_MIN_QT <= q <= ENGINE_OIL_CAPACITY_MAX_QT:
                            candidates.append({
                                "engine": "unknown_engine",
                                "field": "with_filter",
                                "capacity": {"quarts": q, "liters": l},
                                "score": score_capacity_candidate("engine oil with filter " + look_line, target_field="with_filter", engine_key="unknown_engine") + 6,
                            })

            if "engine oil" in line_lower and "with filter" not in line_lower and "capacity" not in line_lower:
                continue

        for m in re.finditer(CAPACITY_PATTERN, text):
            if not is_real_capacity_match(text, m):
                continue
            context = text[max(0, m.start() - 90):min(len(text), m.end() + 90)]
            if "adding engine oil" in context or "dipstick" in context:
                continue
            if any(term in context for term in ENGINE_OIL_STOP_TERMS) and not re.search(
                r"engine\s+oil.{0,50}(?:with\s+filter|filter\s+change|capacity)",
                context,
                re.I,
            ):
                continue
            q, l = to_quarts_liters(m.group(1), m.group(2))
            if ENGINE_OIL_CAPACITY_MIN_QT <= q <= ENGINE_OIL_CAPACITY_MAX_QT:
                field = detect_capacity_field(context) or "with_filter"
                candidates.append({
                    "engine": "unknown_engine",
                    "field": field,
                    "capacity": {"quarts": q, "liters": l},
                    "score": score_capacity_candidate(context, target_field=field, engine_key="unknown_engine"),
                })

    selected = select_best_capacity_candidates(candidates)
    return selected.get("unknown_engine")


def merge_engine_cap_maps(primary, secondary):
    """Merge capacity maps, preferring existing entries and filling missing engines/fields."""
    if not primary:
        return secondary or {}
    if not secondary:
        return primary

    merged = {
        engine: {
            "with_filter": dict(values.get("with_filter")) if values.get("with_filter") else None,
            "without_filter": dict(values.get("without_filter")) if values.get("without_filter") else None,
        }
        for engine, values in primary.items()
    }

    for engine, values in secondary.items():
        current = merged.setdefault(engine, {"with_filter": None, "without_filter": None})
        for field in ("with_filter", "without_filter"):
            if current.get(field) is None and values.get(field) is not None:
                current[field] = dict(values[field])

    return merged

def extract_oils(text):
    """
    Extract oil viscosities, recommendation strength, temperature conditions, and inline
    engine links.
    """
    text = normalize_ocr_oil_text(text)
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
        lowered_window = window.lower()
        boundary_phrases = [
            " as shown in the chart",
            " however,",
            " when it's",
            " when it’s",
            " these numbers",
            " if you are in an area",
        ]
        boundary_positions = [lowered_window.find(phrase) for phrase in boundary_phrases if lowered_window.find(phrase) != -1]
        if boundary_positions:
            window = window[:min(boundary_positions)]
        
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
                        oil_scores[oil] += max(4, score_oil_evidence(line_lower, oil))
                        if eng_size not in engine_oil_map_inline:
                            engine_oil_map_inline[eng_size] = []
                        if oil not in engine_oil_map_inline[eng_size]:
                            engine_oil_map_inline[eng_size].append(oil)
                        temps = extract_temperature(line)
                        final_temps = get_temperature_with_fallback(temps, oil)
                        oil_temps[oil].update(final_temps)

    # Pass 1.5: material/specification tables often wrap "Recommended motor oil"
    # and the Motorcraft SAE grade onto the following line.
    for spec_match in re.finditer(r"(?:recommended|optional)\s+motor\s+oil", text, re.I):
        line_lower = spec_match.group().lower()
        window = text[spec_match.start():min(len(text), spec_match.end() + 220)]
        oils = extract_oil_types_from_text(window)
        if not oils:
            continue

        score_boost = 6 if "recommended motor oil" in line_lower else 2
        for oil in oils:
            if oil in do_not_use_oils:
                continue
            if oil not in oil_scores:
                oil_scores[oil] = 0
                oil_temps[oil] = set()
            oil_scores[oil] += max(score_boost, score_oil_evidence("engine oil " + window, oil))
            oil_temps[oil].add("all temperatures")
    
    # Pass 2: "best viscosity grade" is the strongest primary-oil signal.
    best_oil_pattern = r'sae\s+' + OIL_PATTERN + r'(?:\s+is\s+(?:the\s+)?best|is\s+(?:the\s+)?best\s+viscosity)'
    for match in re.finditer(best_oil_pattern, text.lower(), re.I):
        if match:
            base, grade = match.groups()[-2:]
            oil = normalize_oil(f"{base}W-{grade}")
            if oil not in do_not_use_oils:
                sentence_start = match.start()
                sentence_end = min(len(text), match.end() + 300)
                if oil not in oil_scores:
                    oil_scores[oil] = 0
                    oil_temps[oil] = set()
                oil_scores[oil] += max(20, score_oil_evidence(text[sentence_start:sentence_end], oil))
                
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

            oil_scores[oil] += max(6, score_oil_evidence(context, oil))
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
            
            oil_scores[oil] += max(2, score_oil_evidence(context_lower, oil))
            
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

    # Pass 2.75: support manuals that specify viscosity classes such as
    # "SAE 0W-X or SAE 5W-X, where X stands for 30, 40 or 50."
    for candidate in extract_wildcard_oil_candidates(text):
        oil = candidate["oil"]
        if oil in do_not_use_oils:
            continue

        window = candidate["window"]
        wildcard_engine_context = (
            has_engine_oil_context(window)
            or "sae classes" in window
            or "viscosity ratings" in window
            or "ambient temperatures" in window
        )
        if not wildcard_engine_context:
            continue

        if oil not in oil_scores:
            oil_scores[oil] = 0
            oil_temps[oil] = set()

        if candidate["base"] in {"0W", "5W"}:
            oil_scores[oil] += 4
            oil_temps[oil].add("all temperatures")
        else:
            oil_scores[oil] += 2
            final_temps = get_temperature_with_fallback(extract_temperature(window), oil)
            oil_temps[oil].update(final_temps)

    # Pass 2.8: explicit oil lists such as "SAE 0W-30, SAE 5W-30 or SAE 5W-40".
    for candidate in extract_listed_oil_candidates(text):
        window = candidate["window"]
        window_lower = window.lower()
        if has_non_engine_oil_context(window_lower) or not has_engine_oil_context(window_lower):
            continue

        substitute_only = any(phrase in window_lower for phrase in [
            "if the approved engine oils are not available",
            "if you need to add oil",
            "you may add",
            "up to 1 us quart",
            "no more than 0.5 qt",
            "no more than 0.5 l",
        ])

        for oil in candidate["oils"]:
            if oil in do_not_use_oils:
                continue
            if oil not in oil_scores:
                oil_scores[oil] = 0
                oil_temps[oil] = set()

            score = max(2, score_oil_evidence(window_lower, oil))
            oil_scores[oil] += min(score, 3) if substitute_only else max(score, 5)
            final_temps = get_temperature_with_fallback(extract_temperature(window), oil)
            oil_temps[oil].update(final_temps)
    
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
            
            if "best" in lower:
                oil_scores[oil] += 8
            elif (
                "may be used" in lower
                or "can be used" in lower
                or "can use" in lower
                or "you can use" in lower
                or "consider using" in lower
                or "acceptable" in lower
                or "very cold" in lower
            ):
                oil_scores[oil] += 2
            elif "preferred" in lower or "recommended" in lower:
                oil_scores[oil] += 5
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
                    oil_scores[oil] = max(1, score_oil_evidence(context, oil))
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
    """Run the full PDF-to-JSON extraction pipeline."""
    start_time = time.time()
    source = choose_pdf_source()
    service = None

    match source:
        case "drive":
            service = get_drive_service()
            pdfs = get_all_pdfs(service, FOLDER_ID)
            source_label = "Google Drive"
        case "local":
            pdfs = get_local_pdfs()
            source_label = f"local {LOCAL_MANUALS_FOLDER} folder"
        case _:
            service = get_drive_service()
            pdfs = get_all_pdfs(service, FOLDER_ID)
            source_label = "Google Drive"

    print(f"\nTotal PDFs found in {source_label}: {len(pdfs)}\n")
    results = {}

    # Each PDF is processed independently so partial failures do not block
    # extraction for the rest of the manual set.
    for file in pdfs:
        filename = file["name"]
        print("Processing:", filename)

        year, make, model = parse_filename(filename)

        match source:
            case "local":
                pdf_stream = load_local_pdf(file["path"])
            case _:
                pdf_stream = download_pdf(service, file["id"])

        with fitz.open(stream=pdf_stream.read(), filetype="pdf") as doc:
            extraction_type, avg_chars = analyze_pdf_type(doc)
            
            if extraction_type == "AUTO":
                print(f"  Using text extraction ({avg_chars} chars/page - text-heavy PDF)")
            else:
                print(f"  Using OCR extraction ({avg_chars} chars/page - scanned/manual document)")
            
            # Filename metadata is preferred; the first-page detector only fills
            # gaps for generic file names.
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

            targeted_spec_ocr_text = extract_targeted_spec_ocr_text(doc)
            if targeted_spec_ocr_text.strip():
                raw_text = raw_text + "\n" + targeted_spec_ocr_text
                text = clean_text(raw_text)
                print(f"      Targeted spec OCR extracted {len(targeted_spec_ocr_text)} characters")
            
            # Engine evidence priority: structured tables first, prose fallback second.
            table_engines = extract_engines_from_spec_table(raw_text)
            if table_engines:
                print(f"      TABLE EXTRACTION found engines: {table_engines}")
            
            general_engines = []
            if table_engines:
                all_engines = table_engines
            else:
                general_engines = extract_engines(text)
                all_engines = general_engines
                if general_engines:
                    print(f"      General extraction found engines: {general_engines}")

            if not all_engines:
                variant_engine_labels = extract_variant_engine_labels_from_pdf(doc, make=make, model=model)
                if variant_engine_labels:
                    all_engines = variant_engine_labels
                    print(f"      Variant extraction found engines: {all_engines}")
            
            all_engines = filter_engine_outliers(all_engines)
            all_engines = consolidate_engine_variants(all_engines)
            
            # Capacity extraction merges multiple table strategies from strict to
            # broader parsers so noisy manuals still yield a usable result.
            engine_caps = extract_engine_capacities(doc)
            engine_caps = merge_engine_cap_maps(
                engine_caps,
                extract_model_named_capacity_tables(doc, detected_engines=all_engines, model=model),
            )
            engine_caps = merge_engine_cap_maps(engine_caps, extract_columnar_model_capacity_table(doc))
            shared_fallback_cap = extract_fallback_capacity(doc)
            engine_caps = prefer_shared_capacity_if_current_caps_are_noise(engine_caps, shared_fallback_cap)
            if table_engines and engine_caps:
                general_engines = extract_engines(text)
                supplemented_engines = add_capacity_backed_engine_candidates(
                    all_engines,
                    engine_caps,
                    general_engines,
                )
                if len(supplemented_engines) > len(all_engines):
                    added_engines = supplemented_engines[len(all_engines):]
                    print(f"      Capacity-backed engine supplement found: {added_engines}")
                all_engines = supplemented_engines
            if all_engines and engine_caps:
                engine_caps = filter_engine_caps_to_detected_engines(engine_caps, all_engines)
            
            # When oil capacities are available, they validate which detected
            # engines belong in the final output.
            if all_engines and engine_caps:
                valid_engine_bases = set()
                valid_engine_bodyless_bases = set()
                valid_engine_displacements = set()
                for cap_eng in engine_caps.keys():
                    valid_engine_bases.add(engine_identity_key(cap_eng))
                    valid_engine_bodyless_bases.add(engine_identity_key(cap_eng, strip_body_style=True))
                    displacement = get_engine_displacement(cap_eng)
                    if is_plausible_engine_displacement(displacement):
                        valid_engine_displacements.add(round(displacement, 1))
                
                filtered_engines = []
                for eng in all_engines:
                    eng_base = engine_identity_key(eng)
                    eng_bodyless_base = engine_identity_key(eng, strip_body_style=True)
                    displacement = get_engine_displacement(eng)
                    if (
                        eng_base in valid_engine_bases
                        or eng_bodyless_base in valid_engine_bodyless_bases
                        or (
                            is_plausible_engine_displacement(displacement)
                            and round(displacement, 1) in valid_engine_displacements
                        )
                    ):
                        filtered_engines.append(eng)
                
                if filtered_engines:
                    all_engines = filtered_engines

            # Engine-specific oil rows can add missing engines. Shared-capacity
            # rows stay as unknown_engine until expanded to already detected engines.
            if engine_caps:
                known_engine_bases = {engine_identity_key(eng) for eng in all_engines}
                known_engine_bodyless_bases = {engine_identity_key(eng, strip_body_style=True) for eng in all_engines}
                for cap_eng in engine_caps.keys():
                    if cap_eng == "unknown_engine":
                        continue
                    cap_base = engine_identity_key(cap_eng)
                    cap_bodyless_base = engine_identity_key(cap_eng, strip_body_style=True)
                    if cap_base not in known_engine_bases and cap_bodyless_base not in known_engine_bodyless_bases:
                        all_engines.append(cap_eng)
                        known_engine_bases.add(cap_base)
                        known_engine_bodyless_bases.add(cap_bodyless_base)
            
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
                        normalized_part = normalize_engine_type_token(part)
                        if (
                            normalized_part
                            and is_layout_engine_type(normalized_part)
                            and normalized_part not in engine_types
                        ):
                            engine_types.append(normalized_part)
            
            engine_types_from_text = extract_engine_types(text, all_engines)
            for et in engine_types_from_text:
                if et not in engine_types:
                    engine_types.append(et)
            engine_types = filter_engine_types_by_detected_engines(engine_types, all_engines)
            
            # Oil-to-engine links come from three sources; later merges prioritize
            # section-level and inline evidence over loose proximity matches.
            engine_oil_map = map_oils_to_engines(text)
            engine_oil_map_sections = extract_engine_specific_oil_map(doc)
            engine_oil_map_models = extract_model_specific_oil_map(doc, detected_engines=all_engines, model=model)

            for eng_size, oils_list in engine_oil_map_sections.items():
                if eng_size not in engine_oil_map:
                    engine_oil_map[eng_size] = oils_list[:]
                else:
                    for oil in reversed(oils_list):
                        if oil in engine_oil_map[eng_size]:
                            engine_oil_map[eng_size].remove(oil)
                        engine_oil_map[eng_size].insert(0, oil)

            for eng_size, oils_list in engine_oil_map_models.items():
                if eng_size not in engine_oil_map:
                    engine_oil_map[eng_size] = oils_list[:]
                else:
                    for oil in reversed(oils_list):
                        if oil in engine_oil_map[eng_size]:
                            engine_oil_map[eng_size].remove(oil)
                        engine_oil_map[eng_size].insert(0, oil)
            
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
            if not engine_caps and all_engines and not has_external_engine_oil_capacity_reference(text):
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
                    valid_engine_bases = {engine_identity_key(eng) for eng in all_engines}
                    valid_engine_bodyless_bases = {engine_identity_key(eng, strip_body_style=True) for eng in all_engines}
                    filtered_multi_engine_data = {
                        eng_key: eng_val
                        for eng_key, eng_val in multi_engine_data.items()
                        if (
                            engine_identity_key(eng_key) in valid_engine_bases
                            or engine_identity_key(eng_key, strip_body_style=True) in valid_engine_bodyless_bases
                        )
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
                        temps = sorted(get_temperature_with_fallback(oil_temps.get(oil, []), oil))
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
                                "temperature_condition": temps
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

                elif oil_scores:
                    multi_engine_data = build_oil_only_engine_data(all_engines, oil_scores, oil_temps)

                multi_engine_data = add_missing_engine_type_to_keys(multi_engine_data, engine_types)
            
            # Some manuals contain multiple "model only" targets in one file.
            # Emit one output entry per resolved target model when detected.
            vehicle_targets = build_vehicle_output_targets(filename, raw_text, year, make, model)
            for target in vehicle_targets:
                target_year = target.get("year")
                target_make = target.get("make")
                target_model = target.get("model")
                results[target["result_key"]] = {
                    "Vehicle": {
                        "year": target_year,
                        "make": target_make,
                        "model": target_model,
                        "engine_types": list(engine_types),
                        "displayName": f"{target_year} {target_make} {target_model}"
                    },
                    "engines": multi_engine_data
                }

    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(results, f, indent=4, ensure_ascii=False)

    total_time = time.time() - start_time
    print(f"Total time to read all manuals: {total_time:.2f} seconds")
    print("\nExtraction Complete\n")


if __name__ == "__main__":
    extract_all()
