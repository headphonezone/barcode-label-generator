import re
import openpyxl

ABBR = {
    "headphone": "HP", "zone": "Zone", "earpiece": "EarPc", "eartips": "Tips",
    "cable": "Cbl", "accessory": "Acc", "accessories": "Acc", "version": "Ver",
    "silver": "Slv", "plated": "Plt", "driver": "Drv", "drivers": "Drv",
    "display": "Disp", "leather": "Lthr", "balanced": "Bal", "bal": "Bal",
    "type": "Typ", "pack": "Pk", "wonderwall": "WndrWall", "highway": "Hwy",
    "stairway": "Stway", "everlong": "EvrLng", "protective": "Protc",
    "board": "Bd", "transmission": "Transm",
}
FILLER_WORDS = {"hpz", "fiio", "headphone", "zone", "-", "the", "for", "of"}

# tag patterns: (regex, canonical short tag)
TAG_PATTERNS = [
    (re.compile(r"\[?\bold version\]?", re.I), "OLD"),
    (re.compile(r"\[?\bnew version\]?", re.I), "NEW"),
    (re.compile(r"\bleft\b", re.I), "L"),
    (re.compile(r"\bright\b", re.I), "R"),
]

MODEL_RE = re.compile(r"\b([A-Z]{1,5}\d{1,4}[A-Z]{0,3})\b")
CATEGORY_MAP = [
    (r"cable", 1), (r"earpiece|ear\s?piece|eartip", 2), (r"display", 3),
    (r"case", 4), (r"driver", 5), (r"tube", 6),
]


def extract_tags(text):
    found = []
    stripped = text
    for pat, tag in TAG_PATTERNS:
        if pat.search(stripped):
            found.append(tag)
            stripped = pat.sub(" ", stripped)
    return stripped, found


def clean_base(text):
    text = re.sub(r"\((?:FiiO|Fiio|fiio)\)\s*$", "", text).strip()
    text = re.sub(r"\bhpz\s*x\s*fiio\b", "", text, flags=re.I)
    text = re.sub(r"[\[\]\(\)]", " ", text)
    text = re.sub(r"[\-_/]+", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def drop_fillers(text):
    words = [w for w in text.split(" ") if w.lower() not in FILLER_WORDS]
    return " ".join(words) if words else text  # never fully empty


def abbreviate_words(text):
    words = text.split(" ")
    return " ".join(ABBR.get(w.lower(), w) for w in words)


def build_short_name(original, max_len=26):
    stripped, tags = extract_tags(original)
    base = clean_base(stripped)
    base = drop_fillers(base)
    base = abbreviate_words(base)
    base = re.sub(r"\s+", " ", base).strip()

    tag_suffix = (" " + " ".join(tags)) if tags else ""
    budget = max_len - len(tag_suffix)

    if len(base) > budget:
        base = base[: max(1, budget - 1)].rstrip() + "\u2026"
    short = (base + tag_suffix).strip()
    return short[:max_len]


# ---- Short names for the 21L MRP label ----
# Follows the style of the hand-written short names already in
# "SKU Barcode.xlsx": keep brand / model / variant / colour, drop the
# product-category words, "Unboxed" -> "UB", "With Mic" -> "Mic", etc.
# No length cap -- the 21L label is wide enough and the PDF builder
# shrinks the font to fit.
CATEGORY_SEGMENT_RE = re.compile(
    r"^(?:(?:in|on|over)\s*-?\s*ear\s+)?(?:headphones?|earphones?)$"
    r"|^headphone\s+(?:amp\s*&\s*dac|amplifier|earpads?)$"
    r"|^(?:amp\s*&\s*dac|dac\s*&\s*amp|(?:digital\s+)?audio\s+player|accessory|accessories"
    r"|wireless\s+streaming\s+module|\d+\s+months?\s+warranty)$",
    re.I,
)
LABEL_SHORT_REPLACEMENTS = [
    (re.compile(r"\bwith\s*-?\s*mic\b", re.I), "Mic"),
    (re.compile(r"\bwithout\s*-?\s*mic\b", re.I), "No Mic"),
    (re.compile(r"\b(\d(?:\.\d)?)\s*mm\b", re.I), r"\1"),
    (re.compile(r"\btype\s*-?\s*c\b", re.I), "C"),
    (re.compile(r"\bheadphone\s+zone\b", re.I), "HZ"),
]
PREFIX_RE = re.compile(r"^\s*(unboxed|refurbished)\s*-?\s*", re.I)
PREFIX_SHORT = {"unboxed": "UB", "refurbished": "RF"}


def build_label_short_name(item):
    text = re.sub(r"\s+", " ", str(item)).strip()

    prefix = ""
    m = PREFIX_RE.match(text)
    if m:
        prefix = PREFIX_SHORT[m.group(1).lower()]
        text = text[m.end():]

    segments = [s.strip() for s in re.split(r"\s+-\s+|\s+-(?=\S)|(?<=\S)-\s+", text) if s.strip()]
    kept = [s for s in segments if not CATEGORY_SEGMENT_RE.match(s)] or segments

    short = " ".join(kept)
    for pat, repl in LABEL_SHORT_REPLACEMENTS:
        short = pat.sub(repl, short)
    short = re.sub(r"\s+", " ", f"{prefix} {short}").strip()
    return short


def extract_model_num(original_text):
    m = MODEL_RE.search(original_text.upper())
    if m:
        digits = re.search(r"\d+", m.group(1))
        if digits:
            return int(digits.group(0)) % 1000
    return 0


def extract_category(original_text):
    low = original_text.lower()
    for pat, code in CATEGORY_MAP:
        if re.search(pat, low):
            return code
    return 0


def generate(rows):
    results = []
    used_names = set()
    used_barcodes = set()
    seq = 0
    for sku in rows:
        seq += 1
        short = build_short_name(sku)
        final_short = short
        suffix_n = 2
        while final_short.upper() in used_names:
            trimmed = short[: max(0, 26 - len(str(suffix_n)) - 1)]
            final_short = f"{trimmed}{suffix_n}"
            suffix_n += 1
        used_names.add(final_short.upper())

        model_num = extract_model_num(sku)
        cat_code = extract_category(sku)
        barcode_val = f"{model_num:03d}{cat_code}{seq:04d}"
        while barcode_val in used_barcodes:
            seq += 1
            barcode_val = f"{model_num:03d}{cat_code}{seq:04d}"
        used_barcodes.add(barcode_val)

        results.append((sku, final_short, barcode_val))
    return results


if __name__ == "__main__":
    wb = openpyxl.load_workbook("barcode.xlsx", data_only=True)
    ws = wb["Sheet1"]
    rows = [r[0] for r in ws.iter_rows(min_row=2, values_only=True) if r[0] not in (None, "")]
    print("total rows:", len(rows))
    results = generate(rows)
    for r in results[:25]:
        print(r)
    print("...")
    for r in results[-10:]:
        print(r)
    names = [r[1] for r in results]
    codes = [r[2] for r in results]
    print("unique names:", len(set(names)), "/", len(names))
    print("unique barcodes:", len(set(codes)), "/", len(codes))
    print("max short name len:", max(len(n) for n in names))
