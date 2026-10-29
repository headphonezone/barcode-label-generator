"""
Streamlit UI for the 21L MRP / importer label (barcode on top, importer
details table underneath). Called from app.py when "MRP label" is selected.
"""

import datetime
import re

import pymupdf
import streamlit as st

from mrp_docx_builder import build_mrp_label_docx
from mrp_label_builder import (
    build_mrp_label_pdf, expand_mrp_entries, label_overflows,
    IMPORTER_NAMES, IMPORTER_ADDRESS, IMPORTER_CONTACT, PRODUCT_TYPES, QUANTITY_TEXT,
    LABELS_PER_PAGE, TOP_MARGIN_MM, SIDE_MARGIN_MM, VERTICAL_PITCH_MM,
)

MONTHS = ["JAN", "FEB", "MAR", "APR", "MAY", "JUNE", "JULY", "AUG", "SEPT", "OCT", "NOV", "DEC"]
MAX_LABELS = 1500


def render_page_preview(pdf_data, dpi=110):
    """PNG of the PDF's first page, for the on-screen preview."""
    with pymupdf.open(stream=pdf_data, filetype="pdf") as doc:
        return doc[0].get_pixmap(dpi=dpi).tobytes("png")


# (pattern, product type) -- first match wins, so the more specific come first
PRODUCT_TYPE_GUESSES = [
    (r"cassette", "Stereo Cassette Player"),
    (r"\bcd\b", "CD Audio Player"),
    (r"\bcable\b", "Cable Accessory"),
    (r"audio player|\bdap\b", "Digital Audio Player"),
    (r"speaker", "Speaker"),
    (r"amp\s*(&|/|and)\s*dac", "Headphone Amp/DAC"),
    (r"\bdac\b", "Headphone DAC"),
    (r"amplifier|\bamp\b", "Headphone Amplifier"),
    (r"accessor|eartip|earpad|\bcase\b", "Headphone Accessory"),
    (r"\btws\b|true wireless", "TWS"),
]


def guess_product_type(item):
    """Pre-select the Product dropdown from the item name (still editable).
    None when nothing matches (e.g. plain headphones) -- the user picks."""
    low = item.lower()
    for pattern, product_type in PRODUCT_TYPE_GUESSES:
        if re.search(pattern, low):
            return product_type
    return None


def render_mrp_section(catalog_df):
    if "mrp_items" not in st.session_state:
        st.session_state.mrp_items = []

    st.subheader("2. Build the MRP label")

    # -----------------------------------------------------------------
    # Model name -> barcode (from the SKU catalog)
    # -----------------------------------------------------------------
    if catalog_df is None or len(catalog_df) == 0:
        st.warning("No product catalog loaded — load SKU Barcode.xlsx in step 1.")
        return

    dup_skus = set(catalog_df["sku"][catalog_df["sku"].duplicated(keep=False)])
    # Options are keyed by (item, barcode), not row number: the live sheet
    # can gain rows mid-session, and a row number would then silently point
    # at a different product.
    rows_by_key = {(r["sku"], r["barcode"]): r for _, r in catalog_df.iterrows()}

    def _option_label(k):
        sku, barcode_no = k
        # the same item listed twice has two barcodes -- show which is which
        return f"{sku}  ({barcode_no})" if sku in dup_skus else sku

    picked = st.selectbox(
        "Model name", list(rows_by_key), index=None, format_func=_option_label,
        placeholder="Type to search, e.g. SIMGOT EW300", key="mrp_pick",
    )
    if picked is None:
        st.info("Select a model to fill in its barcode.")
        return

    row = rows_by_key[picked]
    key = "|".join(picked)
    st.text_input("Barcode no. (auto)", value=row["barcode"], disabled=True, key=f"mrp_bc_{key}")
    # not printed on the label -- only names the entry in the print list
    short_name = str(row["label_short_name"])
    # the label prints the product's short name, not the full catalog name
    model_name = st.text_input("Product name on the label (short name)",
                               value=row["label_short_name"], key=f"mrp_model_{key}")
    barcode_value = row["barcode"]

    # -----------------------------------------------------------------
    # Label details (table under the barcode)
    # -----------------------------------------------------------------
    st.markdown("**Label details**")
    importer = st.selectbox("Importer & marketed by", IMPORTER_NAMES)
    st.caption(f"{IMPORTER_ADDRESS}  ·  {IMPORTER_CONTACT}")

    guessed = guess_product_type(row["sku"])
    product_type = st.selectbox("Product", PRODUCT_TYPES,
                                index=PRODUCT_TYPES.index(guessed) if guessed else None,
                                placeholder="Choose the product type",
                                key=f"mrp_type_{key}")

    today = datetime.date.today()
    c3, c4, c5 = st.columns(3)
    with c3:
        month = st.text_input("Month of import", value=f"{MONTHS[today.month - 1]}- {today.year}",
                              placeholder="e.g. SEPT- 2026")
    with c4:
        mrp = st.text_input("MRP (Rs.)", placeholder="e.g. 9999")
    with c5:
        year_of_mfg = st.text_input("Year of Mfg", value=str(today.year))

    manufacturer = st.text_input("Manufacturer's address & country of origin",
                                 placeholder="e.g. LINSOUL INC, NEW YORK")
    st.caption(f"Quantity on label: {QUANTITY_TEXT}")

    qty = st.number_input("Number of labels to print", min_value=1, max_value=MAX_LABELS,
                          value=1, step=1, key="mrp_qty")

    if st.button("Add to list", type="primary", use_container_width=True, key="mrp_add"):
        missing = [name for name, val in (
            ("Barcode", barcode_value), ("Product", product_type or ""), ("Model name", model_name), ("Month of import", month),
            ("MRP", mrp), ("Manufacturer", manufacturer), ("Year of Mfg", year_of_mfg),
        ) if not str(val).strip()]
        if missing:
            st.error(f"Please fill in: {', '.join(missing)}")
        else:
            item = {
                "barcode": barcode_value.strip(), "short_name": short_name.strip(),
                "importer": importer,
                "product_type": product_type, "model_name": model_name.strip(),
                "month_of_import": month.strip(),
                "mrp": mrp.strip().removeprefix("Rs.").removeprefix("Rs").strip(),
                "manufacturer": manufacturer.strip(), "year_of_mfg": year_of_mfg.strip(),
                "qty": int(qty),
            }
            if label_overflows(item):
                st.warning("This text is too long to fit the label at 6 pt — "
                           "shorten the product name or manufacturer address.")
            st.session_state.mrp_items.append(item)
            st.success(f"Added {qty} x {item['short_name']}")

    st.divider()

    # -----------------------------------------------------------------
    # Print list + generation
    # -----------------------------------------------------------------
    st.subheader("3. Print list")
    if not st.session_state.mrp_items:
        st.info("No labels added yet.")
        return

    for idx, item in enumerate(st.session_state.mrp_items):
        c1, c2, c3, c4 = st.columns([4, 3, 2, 1])
        c1.write(f"{item['short_name']}  ·  Rs.{item['mrp']}/-")
        c2.write(item["barcode"])
        c3.write(f"x{item['qty']}")
        if c4.button("🗑️", key=f"mrp_del_{idx}"):
            st.session_state.mrp_items.pop(idx)
            st.rerun()

    total_labels = sum(i["qty"] for i in st.session_state.mrp_items)
    start_position = int(st.number_input(
        "Start at label position on the first sheet",
        min_value=1, max_value=LABELS_PER_PAGE, value=1, step=1, key="mrp_start_pos",
        help="Positions run left→right, top→bottom (1–21). Use this to reuse a partly used sheet.",
    ))
    total_pages = (start_position - 1 + total_labels + LABELS_PER_PAGE - 1) // LABELS_PER_PAGE
    st.markdown(f"**Total labels: {total_labels}**  ·  **Sheets needed: {total_pages}** (21 labels/sheet)")

    with st.expander("⚙️ Alignment fine-tune (only if labels drift on your printed sheet)"):
        tcol1, tcol2, tcol3 = st.columns(3)
        with tcol1:
            top_margin = st.number_input("Top margin (mm)", value=float(TOP_MARGIN_MM), step=0.5,
                                         format="%.1f", key="mrp_top")
        with tcol2:
            side_margin = st.number_input("Side margin (mm)", value=float(SIDE_MARGIN_MM), step=0.5,
                                          format="%.1f", key="mrp_side")
        with tcol3:
            row_pitch = st.number_input("Row pitch (mm)", value=float(VERTICAL_PITCH_MM), step=0.1,
                                        format="%.1f", key="mrp_pitch")
        show_outlines = st.checkbox(
            "Test print: draw label outlines", key="mrp_outlines",
            help="Print on plain A4 paper, lay it over a label sheet against the light and check "
                 "the outlines match the die-cuts. Turn off for the real label sheet.",
        )

    col_a, col_b = st.columns(2)
    with col_a:
        if st.button("Clear all", use_container_width=True, key="mrp_clear"):
            st.session_state.mrp_items = []
            st.rerun()
    with col_b:
        generate = st.button("Generate label sheet", type="primary", use_container_width=True,
                             key="mrp_generate")

    if generate:
        sheet_args = dict(
            start_position=start_position,
            top_margin_mm=top_margin,
            left_margin_mm=side_margin,
            vertical_pitch_mm=row_pitch,
            show_outlines=show_outlines,
        )
        entries = expand_mrp_entries(st.session_state.mrp_items)
        with st.spinner("Building MRP label sheet..."):
            docx_bytes = build_mrp_label_docx(entries, **sheet_args)
            # the PDF is only rendered for the on-screen preview -- exact 21L geometry
            pdf_bytes = build_mrp_label_pdf(entries, **sheet_args).getvalue()
            preview = render_page_preview(pdf_bytes)
        st.success(f"Label sheet ready — {total_labels} labels.")
        st.download_button(
            label="⬇️ Download .docx",
            data=docx_bytes,
            file_name="mrp_labels.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            on_click="ignore",
            use_container_width=True,
        )
        st.image(preview, caption="Page 1 preview (A4)", width="stretch")

    st.divider()
    st.caption("NovaJet MPL 21L — 63.5 x 38 mm, 3 across x 7 down. "
               "From Word, keep scaling at 100% / 'No Scaling' (not 'Scale to paper size'); "
               "if Word warns that margins are outside the printable area, choose Ignore/Yes.")
