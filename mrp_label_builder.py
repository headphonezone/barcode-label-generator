"""
Builds MRP / importer labels on NovaJet MPL 21L sheets
(63.5 x 38 mm, 3 across x 7 down, 21 per page) as a PDF.

Each label has the barcode at the top and, under it, the statutory
importer details table (importer & marketer, product / model / qty /
month of import / MRP, manufacturer & country of origin, year of mfg).

Only as many labels as requested are drawn, so a partly used sheet can
be fed back in -- use start_position to skip labels already peeled off.
"""

import io
from xml.sax.saxutils import escape

from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas
from reportlab.platypus import Paragraph, Table, TableStyle

from label_builder import PAGE_WIDTH_MM, PAGE_HEIGHT_MM, make_barcode_stream

# ---- Label geometry (NovaJet MPL 21L - 63.5 x 38 mm sheet, spec sheet values) ----
TOP_MARGIN_MM = 15.5
SIDE_MARGIN_MM = 7.0
LABEL_WIDTH_MM = 63.5
LABEL_HEIGHT_MM = 38.0
HORIZONTAL_PITCH_MM = 66.0
VERTICAL_PITCH_MM = 38.0
COLS_ACROSS = 3
ROWS_DOWN = 7
LABELS_PER_PAGE = COLS_ACROSS * ROWS_DOWN  # 21

# ---- Layout inside one label (mm, measured from the label's top-left) ----
# Top to bottom: bars, details table (no barcode number). The block is centred
# in the label both ways, so shorter text sits lower instead of hugging the top.
PT_MM = 25.4 / 72
# Margins are tight so all the text fits at 6 pt (checked against the full catalog)
INNER_PAD_MM = 1.0              # left/right
TOP_PAD_MM = 0.5
BOTTOM_PAD_MM = 0.7
BARCODE_GAP_MM = 0.5            # bars -> details table
BARCODE_TOP_MM = TOP_PAD_MM + 0.3
BARCODE_WIDTH_MM = 40.0
BARCODE_HEIGHT_MM = 5.5
TABLE_TOP_MM = BARCODE_TOP_MM + BARCODE_HEIGHT_MM + BARCODE_GAP_MM
LEFT_COL_FRACTION = 0.40        # share of table width for the caption column

# Every letter on the label is Times New Roman 6 pt -- no shrinking. Text
# too long for the label is flagged in the app instead (label_overflows).
FONT_NAME = "Times-Roman"       # reportlab's built-in Times (same metrics as Times New Roman)
LINE_FACTOR = 1.0               # line height / font size -- tight so 6 pt fits the 38 mm label
CELL_PAD_X_MM = 0.5
CELL_PAD_Y_MM = 0.1
FONT_SIZES_PT = [6.0]

# ---- Fixed importer details (same for both importer names) ----
IMPORTER_NAMES = ["Ferrari Video", "Headphone Zone"]
IMPORTER_ADDRESS = "63/18, 1st Main Road, Gandhi Nagar, Chennai-600020"
IMPORTER_CONTACT = "Email: pks@ferrarivideo.com 044-24420733"

PRODUCT_TYPES = ["Headphone", "Amp/DAC", "TWS", "Case"]
QUANTITY_TEXT = "1 unit"

CAPTIONS = [
    "Name and address of Importer &amp; Marketed by &amp; Contact if Consumer Complaints",
    "Product Name / Model ID / Colour / Quantity / Month of Import / MRP (Inclusive of All Taxes)",
    "Manufacturers Address &amp; Country of Origin",
    "Year of Mfg",
]


def _styles(font_pt):
    leading = font_pt * LINE_FACTOR
    left = ParagraphStyle("cap", fontName=FONT_NAME, fontSize=font_pt,
                          leading=leading, alignment=1)
    right = ParagraphStyle("val", fontName=FONT_NAME, fontSize=font_pt,
                           leading=leading, alignment=1)
    return left, right


def _build_table(info, font_pt, width_mm):
    left, right = _styles(font_pt)
    e = lambda s: escape(str(s).strip())

    importer = (f"{e(info['importer'])}, {e(IMPORTER_ADDRESS)}<br/>"
                f"{e(IMPORTER_CONTACT)}")
    product = "<br/>".join([
        e(info["product_type"]),
        e(info["model_name"]),
        e(QUANTITY_TEXT),
        e(info["month_of_import"]),
        f"Rs.{e(info['mrp'])}/-",
    ])
    values = [importer, product, e(info["manufacturer"]), e(info["year_of_mfg"])]

    data = [[Paragraph(cap, left), Paragraph(val, right)]
            for cap, val in zip(CAPTIONS, values)]

    left_w = width_mm * LEFT_COL_FRACTION * mm
    table = Table(data, colWidths=[left_w, width_mm * mm - left_w])
    table.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.3, colors.black),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("BACKGROUND", (0, 0), (0, 0), colors.HexColor("#D9D9D9")),
        ("LEFTPADDING", (0, 0), (-1, -1), CELL_PAD_X_MM * mm),
        ("RIGHTPADDING", (0, 0), (-1, -1), CELL_PAD_X_MM * mm),
        ("TOPPADDING", (0, 0), (-1, -1), CELL_PAD_Y_MM * mm),
        ("BOTTOMPADDING", (0, 0), (-1, -1), CELL_PAD_Y_MM * mm),
    ]))
    return table


def fit_table(info, width_mm, height_mm):
    """Largest font size whose table fits in width x height. Returns (table, height_pt)."""
    for font_pt in FONT_SIZES_PT:
        table = _build_table(info, font_pt, width_mm)
        _, h = table.wrap(width_mm * mm, height_mm * mm)
        if h <= height_mm * mm:
            return table, h
    # Still too tall at the smallest size -- draw it anyway; the caller warns.
    return table, h


def label_overflows(info):
    """True if this label's text can't fit the 21L label even at the smallest font."""
    width, height = _table_box_mm()
    _, h = fit_table(info, width, height)
    return h > height * mm


def _table_box_mm():
    width = LABEL_WIDTH_MM - 2 * INNER_PAD_MM
    height = LABEL_HEIGHT_MM - TABLE_TOP_MM - BOTTOM_PAD_MM
    return width, height


def _draw_label(c, x0, y_top, info, barcode_img):
    width, height = _table_box_mm()
    table, h = fit_table(info, width, height)

    # bars + gap + table as one block, centred vertically in the label
    block_mm = BARCODE_HEIGHT_MM + BARCODE_GAP_MM + h / mm
    block_top = y_top - max(TOP_PAD_MM, (LABEL_HEIGHT_MM - block_mm) / 2) * mm

    # bars, centred horizontally
    bc_x = x0 + (LABEL_WIDTH_MM - BARCODE_WIDTH_MM) / 2 * mm
    bc_y = block_top - BARCODE_HEIGHT_MM * mm
    c.drawImage(barcode_img, bc_x, bc_y, width=BARCODE_WIDTH_MM * mm,
                height=BARCODE_HEIGHT_MM * mm, preserveAspectRatio=False, mask="auto")

    # details table under the bars
    table.drawOn(c, x0 + INNER_PAD_MM * mm, bc_y - BARCODE_GAP_MM * mm - h)


def _draw_outlines(c, top_margin_mm, left_margin_mm, vertical_pitch_mm):
    """Thin grey outline of all 21 die-cut positions -- for test prints on plain paper."""
    c.saveState()
    c.setStrokeColor(colors.HexColor("#9E9E9E"))
    c.setLineWidth(0.25)
    for slot in range(LABELS_PER_PAGE):
        row, col = divmod(slot, COLS_ACROSS)
        x0 = left_margin_mm * mm + col * HORIZONTAL_PITCH_MM * mm
        y_top = (PAGE_HEIGHT_MM - top_margin_mm) * mm - row * vertical_pitch_mm * mm
        c.roundRect(x0, y_top - LABEL_HEIGHT_MM * mm, LABEL_WIDTH_MM * mm,
                    LABEL_HEIGHT_MM * mm, 2 * mm)
    c.restoreState()


def build_mrp_label_pdf(entries, start_position=1, top_margin_mm=TOP_MARGIN_MM,
                        left_margin_mm=SIDE_MARGIN_MM,
                        vertical_pitch_mm=VERTICAL_PITCH_MM,
                        show_outlines=False) -> io.BytesIO:
    """
    entries: list of label-info dicts, one per sticker (quantity already
             expanded). Keys: barcode, short_name, importer, product_type, model_name,
             month_of_import, mrp, manufacturer, year_of_mfg.
    start_position: 1-based label slot on the first page to start printing
             at (for re-using a partly used sheet).
    show_outlines: draw the 21 label outlines (test print on plain paper only).
    Returns an in-memory PDF (BytesIO).
    """
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=(PAGE_WIDTH_MM * mm, PAGE_HEIGHT_MM * mm))

    img_cache = {}
    for info in entries:
        num = info["barcode"]
        if num not in img_cache:
            img_cache[num] = ImageReader(make_barcode_stream(num))

    outline = lambda: show_outlines and _draw_outlines(c, top_margin_mm, left_margin_mm,
                                                       vertical_pitch_mm)
    outline()
    slot = max(1, int(start_position)) - 1
    for info in entries:
        if slot == LABELS_PER_PAGE:
            c.showPage()
            outline()
            slot = 0
        row, col = divmod(slot, COLS_ACROSS)
        x0 = left_margin_mm * mm + col * HORIZONTAL_PITCH_MM * mm
        y_top = (PAGE_HEIGHT_MM - top_margin_mm) * mm - row * vertical_pitch_mm * mm
        _draw_label(c, x0, y_top, info, img_cache[info["barcode"]])
        slot += 1

    c.showPage()
    c.save()
    buf.seek(0)
    return buf


def expand_mrp_entries(items):
    """items: list of label-info dicts each with a 'qty' key -> flat list, one per sticker."""
    expanded = []
    for item in items:
        expanded.extend([item] * int(item["qty"]))
    return expanded
