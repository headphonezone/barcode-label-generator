import datetime

import pandas as pd
import streamlit as st
from label_builder import (
    build_label_sheet, expand_entries, LABELS_PER_PAGE,
    TOP_MARGIN_MM, SIDE_MARGIN_MM, VERTICAL_PITCH_MM,
)
from pdf_label_builder import build_label_sheet_pdf
from catalog_source import (
    load_from_google_sheet, load_from_excel, DEFAULT_CATALOG_PATH, DEFAULT_SHEET_URL,
)
from mrp_app_section import render_mrp_section
from docx.enum.table import WD_ALIGN_VERTICAL

CATALOG_POLL_SECONDS = 10  # how often the live SKU sheet is checked for new products

try:
    SHEET_URL = st.secrets.get("SKU_SHEET_URL", DEFAULT_SHEET_URL)
except Exception:  # no secrets file at all
    SHEET_URL = DEFAULT_SHEET_URL

st.set_page_config(page_title="Barcode Label Generator", page_icon="🏷️", layout="centered")

st.title("🏷️ Barcode Label Generator")
st.caption("NovaJet MPL 21L (MRP labels) and 84L (barcode-only labels)")

if "label_items" not in st.session_state:
    st.session_state.label_items = []
if "catalog_df" not in st.session_state:
    st.session_state.catalog_df = None
    st.session_state.catalog_source = None


@st.cache_data(ttl=CATALOG_POLL_SECONDS, show_spinner=False)
def _load_live_catalog():
    if not SHEET_URL:
        raise ValueError("SKU_SHEET_URL is not set in the app's secrets")
    return load_from_google_sheet(SHEET_URL), datetime.datetime.now()


def _fingerprint(df):
    return int(pd.util.hash_pandas_object(df, index=False).sum())


@st.fragment(run_every=CATALOG_POLL_SECONDS)
def _watch_live_catalog():
    """
    Runs in the background every few seconds. Redraws the whole page only
    when the sheet's contents actually changed, so new products appear in
    the dropdown on their own without the page flickering otherwise.
    """
    if st.session_state.catalog_source != "live":
        return
    try:
        df, _ = _load_live_catalog()
    except Exception:
        return  # network blip -- keep the current list, try again next tick
    if _fingerprint(df) != st.session_state.get("catalog_fp"):
        st.rerun(scope="app")


# ---------------------------------------------------------------------
# Catalog source: the live SKU Barcode Google Sheet by default (watched
# in the background so new products show up automatically), or a
# manually loaded sheet / Excel file.
# ---------------------------------------------------------------------
st.subheader("1. Product catalog")
if st.session_state.catalog_source in (None, "live"):
    try:
        df, loaded_at = _load_live_catalog()
        fp = _fingerprint(df)
        prev_count = len(st.session_state.catalog_df) if st.session_state.catalog_df is not None else None
        if st.session_state.get("catalog_fp") not in (None, fp):
            added = len(df) - (prev_count or 0)
            st.toast(f"SKU sheet updated — {added} new product(s)" if added > 0
                     else "SKU sheet updated", icon="🔄")
        st.session_state.catalog_df, st.session_state.catalog_fp = df, fp
        st.session_state.catalog_source = "live"
        # the sheet link itself isn't shown -- the app may be publicly reachable
        st.caption(f"✅ {len(df)} products from the live SKU Barcode sheet "
                   f"· checked {loaded_at:%H:%M:%S}. New products appear automatically.")
        _watch_live_catalog()
    except Exception as e:
        if DEFAULT_CATALOG_PATH.exists():
            st.session_state.catalog_df = load_from_excel(DEFAULT_CATALOG_PATH)
            st.warning(f"Couldn't reach the Google Sheet ({e}). Using the local copy "
                       f"{DEFAULT_CATALOG_PATH.name} instead — it may be out of date.")
        else:
            st.error(f"Couldn't reach the Google Sheet: {e}")
else:
    st.caption(f"✅ {len(st.session_state.catalog_df)} products from a manually loaded file.")
    if st.button("Switch back to the live SKU Barcode sheet"):
        st.session_state.catalog_source = None
        st.rerun()

with st.expander("📄 Load a different catalog (Google Sheet or Excel)", expanded=st.session_state.catalog_df is None):
    st.caption(
        "Required columns: **SKU**, **Barcode**. A **Short name** column is optional — "
        "if it's missing, one is auto-generated from the SKU text. "
        "For Google Sheets: File → Share → \"Anyone with the link\" → Viewer, then paste the link below."
    )
    tab1, tab2 = st.tabs(["Google Sheet", "Upload Excel"])

    with tab1:
        sheet_url = st.text_input("Google Sheet link", placeholder="https://docs.google.com/spreadsheets/d/...")
        if st.button("Load / Refresh from Google Sheet"):
            if not sheet_url.strip():
                st.error("Paste a Google Sheet link first.")
            else:
                try:
                    with st.spinner("Loading sheet..."):
                        st.session_state.catalog_df = load_from_google_sheet(sheet_url)
                    st.session_state.catalog_source = "manual"
                    st.success(f"Loaded {len(st.session_state.catalog_df)} products.")
                except Exception as e:
                    st.error(f"Couldn't load that sheet: {e}")

    with tab2:
        uploaded = st.file_uploader("Excel file (.xlsx)", type=["xlsx"])
        if uploaded is not None:
            try:
                st.session_state.catalog_df = load_from_excel(uploaded)
                st.session_state.catalog_source = "manual"
                st.success(f"Loaded {len(st.session_state.catalog_df)} products.")
            except Exception as e:
                st.error(f"Couldn't read that file: {e}")

catalog_df = st.session_state.catalog_df

st.divider()

label_type = st.radio(
    "Label type",
    ["MRP label (21L — 63.5 x 38 mm)", "Barcode only (84L — 46 x 11 mm)"],
    horizontal=True,
)
if label_type.startswith("MRP"):
    render_mrp_section(catalog_df)
    st.stop()

# ---------------------------------------------------------------------
# Add a product: pick from catalog (autofills short name + barcode), or
# enter manually if it's not in the sheet yet.
# ---------------------------------------------------------------------
st.subheader("2. Add a product to the print list")

if catalog_df is not None and len(catalog_df) > 0:
    options = [""] + [
        f"{row.sku}  →  {row.short_name}  ({row.barcode})" for row in catalog_df.itertuples()
    ]
    picked = st.selectbox("Search / select a product", options, index=0)

    if picked:
        picked_idx = options.index(picked) - 1
        row = catalog_df.iloc[picked_idx]
        c1, c2 = st.columns(2)
        with c1:
            st.text_input("Short name (from catalog)", value=row["short_name"], disabled=True)
        with c2:
            st.text_input("Barcode (from catalog)", value=row["barcode"], disabled=True)

        qty = st.number_input("Number of stickers to print", min_value=1, max_value=5000, value=1, step=1, key="qty_catalog")
        if st.button("Add to list", type="primary", use_container_width=True, key="add_catalog"):
            st.session_state.label_items.append({
                "name": row["short_name"],
                "barcode": row["barcode"],
                "qty": int(qty),
            })
            st.success(f"Added {qty} x {row['short_name']}")

    with st.expander("Not in the catalog? Add manually"):
        with st.form("manual_add_form", clear_on_submit=True):
            mcol1, mcol2 = st.columns(2)
            with mcol1:
                m_name = st.text_input("Product / model name", placeholder="e.g. FIIO JD10 C Mic Black")
            with mcol2:
                m_barcode = st.text_input("Barcode value", placeholder="e.g. 54654168463")
            m_qty = st.number_input("Number of stickers to print", min_value=1, max_value=5000, value=1, step=1)
            m_submitted = st.form_submit_button("Add to list", use_container_width=True)
            if m_submitted:
                if not m_name.strip():
                    st.error("Please enter a product name.")
                elif not m_barcode.strip():
                    st.error("Please enter a barcode value.")
                else:
                    st.session_state.label_items.append({
                        "name": m_name.strip(), "barcode": m_barcode.strip(), "qty": int(m_qty),
                    })
                    st.success(f"Added {m_qty} x {m_name.strip()}")
else:
    st.info("Connect a catalog above to search & autofill, or add a product manually below.")
    with st.form("manual_add_form_nocat", clear_on_submit=True):
        mcol1, mcol2 = st.columns(2)
        with mcol1:
            m_name = st.text_input("Product / model name", placeholder="e.g. FIIO JD10 C Mic Black")
        with mcol2:
            m_barcode = st.text_input("Barcode value", placeholder="e.g. 54654168463")
        m_qty = st.number_input("Number of stickers to print", min_value=1, max_value=5000, value=1, step=1)
        m_submitted = st.form_submit_button("Add to list", use_container_width=True)
        if m_submitted:
            if not m_name.strip():
                st.error("Please enter a product name.")
            elif not m_barcode.strip():
                st.error("Please enter a barcode value.")
            else:
                st.session_state.label_items.append({
                    "name": m_name.strip(), "barcode": m_barcode.strip(), "qty": int(m_qty),
                })
                st.success(f"Added {m_qty} x {m_name.strip()}")

st.divider()

# ---------------------------------------------------------------------
# Print list + generation
# ---------------------------------------------------------------------
st.subheader("3. Print list")
if not st.session_state.label_items:
    st.info("No products added yet.")
else:
    total_labels = sum(i["qty"] for i in st.session_state.label_items)
    total_pages = (total_labels + LABELS_PER_PAGE - 1) // LABELS_PER_PAGE

    for idx, item in enumerate(st.session_state.label_items):
        c1, c2, c3, c4 = st.columns([4, 3, 2, 1])
        c1.write(item["name"])
        c2.write(item["barcode"])
        c3.write(f"x{item['qty']}")
        if c4.button("🗑️", key=f"del_{idx}"):
            st.session_state.label_items.pop(idx)
            st.rerun()

    st.markdown(f"**Total stickers: {total_labels}**  ·  **Pages needed: {total_pages}** (84 labels/page)")

    output_format = st.radio(
        "Output format",
        ["PDF (recommended)", "Word (.docx)"],
        horizontal=True,
        help="PDF draws every label at a fixed coordinate, so it prints identically everywhere. "
             "Word's layout can be interpreted slightly differently by different Word versions/OSes "
             "(this caused the earlier alignment drift) — use it only if you need to edit the file in Word.",
    )

    with st.expander("⚙️ Alignment fine-tune (only if labels drift on your printed sheet)"):
        st.caption(
            "Content is top-aligned inside each label automatically. "
            "If the whole grid still sits slightly off the die-cut lines on your "
            "printer, nudge these and regenerate — print one test page first."
        )
        tcol1, tcol2, tcol3 = st.columns(3)
        with tcol1:
            top_margin = st.number_input("Top margin (mm)", value=float(TOP_MARGIN_MM), step=0.5, format="%.1f")
        with tcol2:
            side_margin = st.number_input("Side margin (mm)", value=float(SIDE_MARGIN_MM), step=0.5, format="%.1f")
        with tcol3:
            row_pitch = st.number_input(
                "Row pitch (mm)", value=float(VERTICAL_PITCH_MM), step=0.1, format="%.1f",
                help="Top-of-label to top-of-next-label spacing, from your sheet's spec. "
                     "The gap between rows is computed automatically as pitch minus label height."
            )

        if output_format == "Word (.docx)":
            v_position = st.selectbox(
                "Content position within each label",
                ["Top", "Center", "Bottom"],
                index=0,
                help="If there's blank space above your printed content, use Top. "
                     "If it looks squeezed to one side, try Center.",
            )
            v_align_map = {"Top": WD_ALIGN_VERTICAL.TOP, "Center": WD_ALIGN_VERTICAL.CENTER, "Bottom": WD_ALIGN_VERTICAL.BOTTOM}

    col_a, col_b = st.columns(2)
    with col_a:
        if st.button("Clear all", use_container_width=True):
            st.session_state.label_items = []
            st.rerun()
    with col_b:
        generate = st.button("Generate label sheet", type="primary", use_container_width=True)

    if generate:
        entries = expand_entries(st.session_state.label_items)
        if output_format == "PDF (recommended)":
            with st.spinner("Building label sheet (PDF)..."):
                pdf_bytes = build_label_sheet_pdf(
                    entries,
                    top_margin_mm=top_margin,
                    left_margin_mm=side_margin,
                    vertical_pitch_mm=row_pitch,
                )
            st.success("Label sheet ready.")
            st.download_button(
                label="⬇️ Download .pdf",
                data=pdf_bytes,
                file_name="barcode_labels.pdf",
                mime="application/pdf",
                use_container_width=True,
            )
        else:
            with st.spinner("Building label sheet (Word)..."):
                docx_bytes = build_label_sheet(
                    entries,
                    top_margin_mm=top_margin,
                    left_margin_mm=side_margin,
                    vertical_pitch_mm=row_pitch,
                    vertical_align=v_align_map[v_position],
                )
            st.success("Label sheet ready.")
            st.download_button(
                label="⬇️ Download .docx",
                data=docx_bytes,
                file_name="barcode_labels.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
            )

st.divider()
st.caption(
    "Labels are sized to fit the 46 x 11 mm frame automatically — "
    "no manual resizing needed in Word. Print at 100% scale (no 'fit to page')."
)
