"""
Builds the same NovaJet MPL 84L label sheet as label_builder.py, but as a
PDF instead of a .docx.

Why this exists: a .docx's layout is only a set of *instructions* --
Word (and different Word versions/OSes) can interpret row heights, empty
paragraph line heights, and vertical alignment slightly differently, which
is exactly what caused the earlier gap-row drift bug. A PDF has no such
ambiguity: every piece of content is drawn at an explicit (x, y) coordinate
on a fixed-size page. What you generate is pixel-for-pixel what prints,
regardless of what computer or printer driver opens it.
"""

import io
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader

from label_builder import (
    PAGE_WIDTH_MM, PAGE_HEIGHT_MM, TOP_MARGIN_MM, SIDE_MARGIN_MM,
    LABEL_WIDTH_MM, LABEL_HEIGHT_MM, HORIZONTAL_PITCH_MM, VERTICAL_PITCH_MM,
    COLS_ACROSS, ROWS_DOWN, LABELS_PER_PAGE,
    BARCODE_WIDTH_MM, BARCODE_HEIGHT_MM, NAME_FONT_PT,
    make_barcode_stream,
)

# Vertical layout within a single label (all "from the top of the label"):
TEXT_BASELINE_OFFSET_MM = 2.3   # distance from label top to text baseline
BARCODE_TOP_GAP_MM = 2.6        # distance from label top to barcode's top edge


def make_barcode_image_reader(number: str) -> ImageReader:
    return ImageReader(make_barcode_stream(number))


def build_label_sheet_pdf(entries, top_margin_mm=TOP_MARGIN_MM, left_margin_mm=SIDE_MARGIN_MM,
                           vertical_pitch_mm=VERTICAL_PITCH_MM) -> io.BytesIO:
    """
    entries: list of (name, barcode_value) tuples, one per sticker
             (quantity already expanded).
    Returns an in-memory PDF (BytesIO), one page per 84 labels.
    """
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=(PAGE_WIDTH_MM * mm, PAGE_HEIGHT_MM * mm))

    img_cache = {}
    for _, num in entries:
        if num not in img_cache:
            img_cache[num] = make_barcode_image_reader(num)

    total = len(entries)
    idx = 0

    def draw_page(page_entries):
        for pos, (name, num) in enumerate(page_entries):
            row = pos // COLS_ACROSS
            col = pos % COLS_ACROSS

            x0 = left_margin_mm * mm + col * HORIZONTAL_PITCH_MM * mm
            y_top = (PAGE_HEIGHT_MM - top_margin_mm) * mm - row * vertical_pitch_mm * mm

            # product name, centered, top of label
            c.setFont("Helvetica-Bold", NAME_FONT_PT)
            text_y = y_top - TEXT_BASELINE_OFFSET_MM * mm
            c.drawCentredString(x0 + LABEL_WIDTH_MM * mm / 2, text_y, name)

            # barcode, centered, below the name
            bc_w = BARCODE_WIDTH_MM * mm
            bc_h = BARCODE_HEIGHT_MM * mm
            bc_x = x0 + (LABEL_WIDTH_MM * mm - bc_w) / 2
            bc_top_y = y_top - BARCODE_TOP_GAP_MM * mm
            bc_y = bc_top_y - bc_h
            c.drawImage(img_cache[num], bc_x, bc_y, width=bc_w, height=bc_h,
                        preserveAspectRatio=False, mask='auto')

    if total == 0:
        c.showPage()
    else:
        while idx < total:
            page_entries = entries[idx: idx + LABELS_PER_PAGE]
            draw_page(page_entries)
            idx += LABELS_PER_PAGE
            c.showPage()

    c.save()
    buf.seek(0)
    return buf
