"""
Core logic for building NovaJet MPL 84L (46 x 11 mm, 4 across x 21 down)
barcode label sheets. Used by app.py (Streamlit UI). Kept dependency-free
of any uploaded template file so it can run anywhere.
"""

import io
import barcode
from barcode.writer import ImageWriter
from PIL import Image
import docx
from docx.shared import Mm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ROW_HEIGHT_RULE
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.oxml.shape import CT_Inline
from docx.shape import InlineShape

# ---- Label geometry (NovaJet MPL 84L - 46 x 11 mm sheet) ----
# All of these come directly off the spec sheet. Nothing here is a guessed
# or "close enough" value -- the two gaps (between columns, between rows)
# are DERIVED from pitch minus label size, exactly as the spec sheet
# defines pitch (top-of-label to top-of-next-label).
PAGE_WIDTH_MM = 210
PAGE_HEIGHT_MM = 297
TOP_MARGIN_MM = 18.0
SIDE_MARGIN_MM = 5.5
LABEL_WIDTH_MM = 46.0
LABEL_HEIGHT_MM = 11.0
HORIZONTAL_PITCH_MM = 51.0   # label-to-label horizontal spacing (spec sheet)
VERTICAL_PITCH_MM = 12.5     # label-to-label vertical spacing (spec sheet)
COLS_ACROSS = 4
ROWS_DOWN = 21
LABELS_PER_PAGE = COLS_ACROSS * ROWS_DOWN  # 84

# computed, not hardcoded
GAP_WIDTH_MM = HORIZONTAL_PITCH_MM - LABEL_WIDTH_MM    # 51.0 - 46.0 = 5.0mm
GAP_HEIGHT_MM = VERTICAL_PITCH_MM - LABEL_HEIGHT_MM    # 12.5 - 11.0 = 1.5mm

BARCODE_WIDTH_MM = 42
BARCODE_HEIGHT_MM = 6.8
NAME_FONT_PT = 6

# Microsoft Word reserves a 0.75 pt table-row boundary between consecutive
# rows, including borderless rows.  Row.Height reports only the requested
# height and does not include that boundary, but printed/PDF output does.
# Subtract it from the requested table row so the rendered top-to-top pitch
# remains the physical 12.5 mm specified by NovaJet.
WORD_ROW_BOUNDARY_PT = 0.75


def make_barcode_stream(number: str) -> io.BytesIO:
    """Render a Code128 barcode (no human-readable text) into an in-memory PNG."""
    code128 = barcode.get('code128', str(number), writer=ImageWriter())
    # 300dpi is already well above what any barcode scanner or printer
    # needs (thermal label printers are typically 200-300dpi) -- 600dpi
    # just bloats file size with detail nobody uses, especially once
    # multiplied across hundreds of unique barcodes in one document.
    options = dict(module_width=0.35, module_height=18, quiet_zone=1.0,
                   write_text=False, dpi=150)
    raw = io.BytesIO()
    code128.write(raw, options)
    raw.seek(0)

    # Re-save as 1-bit (bilevel) -- a barcode is pure black/white, so this
    # is lossless here and shrinks the embedded image substantially
    # compared to the RGB PNG python-barcode produces by default.
    img = Image.open(raw).convert("1")
    out = io.BytesIO()
    img.save(out, format="PNG")
    out.seek(0)
    return out


def _set_cell_margins(cell, top=15, bottom=15, left=20, right=20):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    mar = OxmlElement('w:tcMar')
    for tag, val in (('top', top), ('bottom', bottom), ('start', left), ('end', right)):
        node = OxmlElement(f'w:{tag}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        mar.append(node)
    tcPr.append(mar)


def _set_no_spacing(paragraph):
    pf = paragraph.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing = 1.0
    # Keep Word's document line grid from moving text independently of the
    # fixed label geometry.
    pPr = paragraph._p.get_or_add_pPr()
    snap = pPr.find(qn('w:snapToGrid'))
    if snap is None:
        snap = OxmlElement('w:snapToGrid')
        pPr.append(snap)
    snap.set(qn('w:val'), '0')


def _new_label_table(document, vertical_pitch_mm=VERTICAL_PITCH_MM):
    """
    Add one 84-label table using the same structure as NovaJet's supplied
    84L Word template: 21 rows, each exactly one vertical pitch high.

    The physical label occupies the first 11 mm of each 12.5 mm pitch; the
    remaining 1.5 mm is the gap before the next label.  Keeping both parts
    in one exact-height Word row is important.  Separate 1.5 mm table rows
    can be expanded by Word's hidden paragraph minimums and cause a small
    error to accumulate down the page.
    """
    table = document.add_table(rows=ROWS_DOWN, cols=COLS_ACROSS * 2 - 1)
    table.autofit = False

    # Declare borderless cells explicitly.  Relying on Word's implicit
    # table-grid setting can reserve a small boundary at every row even
    # when no border is visibly printed.
    tbl_pr = table._tbl.tblPr
    borders = OxmlElement('w:tblBorders')
    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        border = OxmlElement(f'w:{edge}')
        border.set(qn('w:val'), 'nil')
        borders.append(border)
    tbl_pr.append(borders)

    # column widths: label, gap, label, gap, label, gap, label
    widths = []
    for i in range(COLS_ACROSS * 2 - 1):
        widths.append(Mm(LABEL_WIDTH_MM) if i % 2 == 0 else Mm(GAP_WIDTH_MM))

    # tblGrid must carry the real widths -- Word/renderers use the grid
    # for column layout and text-wrap decisions, not just the per-cell
    # width set below.
    for i, col in enumerate(table.columns):
        col.width = widths[i]

    for row in table.rows:
        row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
        row.height = Mm(vertical_pitch_mm) - Pt(WORD_ROW_BOUNDARY_PT)
        for i, cell in enumerate(row.cells):
            cell.width = widths[i]

    return table


class _FastPicInserter:
    """
    Drop-in replacement for Run.add_picture() that avoids python-docx's
    Run.add_picture -> next_id property, which does a full-document XPath
    scan (//@id) on every single call. That makes bulk insertion of
    thousands of pictures effectively O(n^2). We keep our own counter
    instead, which is safe because we're building the document from
    scratch and are the only source of id-bearing elements.
    """
    def __init__(self, document):
        self.document = document
        self._next_shape_id = 1

    def add_picture(self, paragraph, image_stream, width, height):
        rId, image = self.document.part.get_or_add_image(image_stream)
        cx, cy = image.scaled_dimensions(width, height)
        shape_id = self._next_shape_id
        self._next_shape_id += 1
        inline = CT_Inline.new_pic_inline(shape_id, rId, image.filename, cx, cy)
        run = paragraph.add_run()
        run._r.add_drawing(inline)
        return InlineShape(inline)


def _fill_label_cell(cell, name, barcode_stream, pic_inserter, vertical_align=WD_ALIGN_VERTICAL.TOP):
    cell.text = ""
    _set_cell_margins(cell)
    cell.vertical_alignment = vertical_align

    p1 = cell.paragraphs[0]
    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _set_no_spacing(p1)
    run1 = p1.add_run(name)
    run1.font.size = Pt(NAME_FONT_PT)
    run1.font.bold = True
    run1.font.name = "Arial"

    p2 = cell.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _set_no_spacing(p2)
    barcode_stream.seek(0)
    pic_inserter.add_picture(p2, barcode_stream, Mm(BARCODE_WIDTH_MM), Mm(BARCODE_HEIGHT_MM))


def build_label_sheet(entries, top_margin_mm=TOP_MARGIN_MM, left_margin_mm=SIDE_MARGIN_MM,
                       vertical_pitch_mm=VERTICAL_PITCH_MM,
                       vertical_align=WD_ALIGN_VERTICAL.TOP) -> io.BytesIO:
    """
    entries: list of (name, barcode_value) tuples, already expanded so each
             appears once per sticker (i.e. quantity already applied).
    top_margin_mm / left_margin_mm / vertical_pitch_mm: nudge these slightly if
    your printed sheet is drifting off the die-cut lines (printer feed
    variance differs machine to machine). vertical_pitch_mm is the distance
    from the top of one label to the top of the next.
    Returns an in-memory .docx file (BytesIO) with as many 84-label pages
    as needed.
    """
    document = docx.Document()
    section = document.sections[0]
    section.page_width = Mm(PAGE_WIDTH_MM)
    section.page_height = Mm(PAGE_HEIGHT_MM)
    section.top_margin = Mm(top_margin_mm)
    section.bottom_margin = Mm(0)
    section.left_margin = Mm(left_margin_mm)
    section.right_margin = Mm(left_margin_mm)

    # cache one barcode image per unique value
    img_cache = {}
    for _, num in entries:
        if num not in img_cache:
            img_cache[num] = make_barcode_stream(num)

    pic_inserter = _FastPicInserter(document)

    total = len(entries)
    idx = 0
    first_page = True

    while idx < total or (first_page and total == 0):
        if not first_page:
            document.add_page_break()
        first_page = False

        table = _new_label_table(document, vertical_pitch_mm=vertical_pitch_mm)
        for row in table.rows:
            for col in range(0, COLS_ACROSS * 2 - 1, 2):
                if idx >= total:
                    continue
                name, num = entries[idx]
                img_stream = img_cache[num]
                img_stream.seek(0)
                _fill_label_cell(row.cells[col], name, img_stream, pic_inserter, vertical_align=vertical_align)
                idx += 1

    out = io.BytesIO()
    document.save(out)
    out.seek(0)
    return out


def expand_entries(items):
    """items: list of dicts with keys name, barcode, qty -> flat list of (name, barcode)."""
    expanded = []
    for item in items:
        for _ in range(int(item["qty"])):
            expanded.append((item["name"], item["barcode"]))
    return expanded
