# app.py
# TFP Dynamic Restock – Streamlit app
# - Minimal settings (lead time, target coverage, safety buffer, rounding/moq, optional year_tag)
# - Excel export with tabs: FULL, ADMIN, SUPPLIER, PO READY, SETTINGS
# - PO READY: Our Code = Color SKU = Variant SKU without last 3 digits (master 5 + color 3)
# - Notes column after Color, Total Units before Photo
# - Excel print settings for PO READY: fit all columns to page width + repeat header row on each page

from __future__ import annotations

import io
import re
from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple

import numpy as np
import pandas as pd
import streamlit as st


# -----------------------------
# Utilities
# -----------------------------

def _norm(s: str) -> str:
    return re.sub(r"\s+", " ", str(s).strip().lower())


def _safe_int(x, default: int = 0) -> int:
    try:
        if pd.isna(x):
            return default
        return int(float(x))
    except Exception:
        return default


def _safe_float(x, default: float = 0.0) -> float:
    try:
        if pd.isna(x):
            return default
        return float(x)
    except Exception:
        return default


def excel_col(n: int) -> str:
    """0-indexed to Excel letters."""
    n += 1
    out = ""
    while n:
        n, r = divmod(n - 1, 26)
        out = chr(65 + r) + out
    return out


def round_up_to_step(x: float, step: int) -> int:
    if step <= 1:
        return int(np.ceil(max(0, x)))
    return int(np.ceil(max(0, x) / step) * step)


def parse_variant_sku_parts(variant_sku: str) -> Tuple[Optional[str], Optional[str], Optional[str]]:
    """
    Variant SKU format:
    - first 5: master code
    - next 3: color code
    - last 3: size code
    Example: 21989 001 042  (master=21989, color=001, size=042)
    Returns (master, color, size) where any can be None if parsing fails.
    """
    if variant_sku is None:
        return None, None, None
    s = re.sub(r"\s+", "", str(variant_sku))
    if len(s) < 11:
        return None, None, None
    master = s[:5]
    color = s[5:8]
    size = s[-3:]
    return master, color, size


def derive_our_code(variant_sku: str) -> Optional[str]:
    """
    Our Code (Color SKU) = variant SKU without last 3 digits.
    i.e. master(5) + color(3) = first 8 characters.
    """
    if variant_sku is None:
        return None
    s = re.sub(r"\s+", "", str(variant_sku))
    if len(s) < 8:
        return None
    return s[:-3] if len(s) >= 11 else s[:8]


# -----------------------------
# Settings
# -----------------------------

@dataclass(frozen=True)
class Settings:
    year_tag: str
    lead_time_weeks: float
    target_weeks: float
    safety_buffer: float
    rounding_step: int
    moq: int
    sales_window_weeks: float


# -----------------------------
# Input loading & normalization
# -----------------------------

COLUMN_SYNONYMS = {
    "supplier_code": [
        "supplier code", "vendor code", "supplier", "vendor",
        "κωδικος προμηθευτη", "κωδικός προμηθευτή", "προμηθευτης",
    ],
    "variant_sku": [
        "variant sku", "sku", "internal code", "εσωτερικος κωδικος", "εσωτερικός κωδικός", "internal sku",
    ],
    "color_name": ["color", "χρωμα", "χρώμα"],
    "product_type": ["product type", "type", "category", "product category"],
    "product_name": ["name", "product name", "title", "product title"],
    "photo": ["photo", "image", "image url", "photo url", "φωτο", "εικόνα", "εικονα"],
    "forecasted": ["forecasted", "forecasted (odoo)", "qty forecasted", "stock forecasted"],
    "incoming": ["incoming", "in transit", "qty incoming"],
    "qty_delivered": ["qty delivered", "delivered", "odoo qty delivered"],
    "on_hand": ["on hand", "qty on hand", "stock on hand", "available"],
    "outgoing": ["outgoing", "reserved", "allocated", "qty outgoing"],
    "units_sold": ["units sold", "sold", "qty sold", "sales units", "units"],
    "season": ["season", "year", "tag", "season tag", "collection"],
    "size": ["size", "νούμερο", "νουμερο", "νούμερα"],
}


def canonicalize_columns(df: pd.DataFrame) -> Tuple[pd.DataFrame, Dict[str, str]]:
    """
    Returns df with same columns, plus a mapping from canonical key -> actual column name (if found).
    """
    col_map: Dict[str, str] = {}
    norm_cols = {c: _norm(c) for c in df.columns}

    for key, syns in COLUMN_SYNONYMS.items():
        for c, nc in norm_cols.items():
            if nc == _norm(key) or any(nc == _norm(s) for s in syns):
                col_map[key] = c
                break

    return df, col_map


def read_table(file) -> pd.DataFrame:
    name = file.name.lower()
    if name.endswith(".csv"):
        return pd.read_csv(file)
    if name.endswith(".xlsx") or name.endswith(".xls"):
        return pd.read_excel(file, sheet_name=0)
    raise ValueError("Unsupported file format. Please upload CSV or XLSX.")


def normalize_input(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    for c in df.columns:
        if df[c].dtype == object:
            df[c] = df[c].astype(str).str.strip()
    return df


# -----------------------------
# Core calculations
# -----------------------------

def compute_engine(df: pd.DataFrame, col_map: Dict[str, str], s: Settings) -> pd.DataFrame:
    df = df.copy()

    if "variant_sku" not in col_map:
        raise ValueError("Could not detect Variant SKU column (e.g., 'Variant SKU' / 'SKU' / 'ΕΣΩΤΕΡΙΚΟΣ ΚΩΔΙΚΟΣ').")

    if s.year_tag.strip():
        season_col = col_map.get("season")
        if season_col and season_col in df.columns:
            df = df[df[season_col].astype(str).str.contains(s.year_tag.strip(), na=False)]

    variant_col = col_map["variant_sku"]

    parts = df[variant_col].apply(parse_variant_sku_parts)
    df["Master Code"] = parts.apply(lambda t: t[0])
    df["Color Code"] = parts.apply(lambda t: t[1])
    df["Size Code"] = parts.apply(lambda t: t[2])
    df["Our Code"] = df[variant_col].apply(derive_our_code)

    size_col = col_map.get("size")
    if size_col and size_col in df.columns:
        df["Size"] = df[size_col].apply(lambda x: _safe_int(x, default=np.nan) if str(x).strip().isdigit() else str(x).strip())
    else:
        df["Size"] = df["Size Code"].apply(lambda x: _safe_int(x, default=np.nan))

    forecasted_col = col_map.get("forecasted")
    incoming_col = col_map.get("incoming")
    on_hand_col = col_map.get("on_hand")
    outgoing_col = col_map.get("outgoing")
    qty_delivered_col = col_map.get("qty_delivered")
    units_sold_col = col_map.get("units_sold")

    def col_or_zero(cname: Optional[str]) -> pd.Series:
        if cname and cname in df.columns:
            return df[cname].apply(_safe_float)
        return pd.Series([0.0] * len(df), index=df.index)

    df["Forecasted (Odoo)"] = col_or_zero(forecasted_col)
    df["Incoming"] = col_or_zero(incoming_col)
    df["On Hand"] = col_or_zero(on_hand_col)
    df["Outgoing/Reserved"] = col_or_zero(outgoing_col)
    df["Odoo Qty Delivered"] = col_or_zero(qty_delivered_col)

    computed_forecasted = df["On Hand"] - df["Outgoing/Reserved"] + df["Incoming"]
    df["Effective Forecasted"] = np.where(df["Forecasted (Odoo)"] != 0, df["Forecasted (Odoo)"], computed_forecasted)

    sold_base = col_or_zero(units_sold_col)
    sold_base = np.where(sold_base != 0, sold_base, df["Odoo Qty Delivered"])
    df["_sold_base"] = sold_base

    window_weeks = max(1e-6, float(s.sales_window_weeks))
    df["Units/week adj"] = df["_sold_base"] / window_weeks

    safety = max(0.0, float(s.safety_buffer))
    df["Demand units (target)"] = df["Units/week adj"] * float(s.target_weeks) * (1.0 + safety)
    df["Gap"] = df["Demand units (target)"] - df["Effective Forecasted"]
    df["Reorder point"] = df["Units/week adj"] * float(s.lead_time_weeks) * (1.0 + safety)

    def suggest(row) -> int:
        gap = float(row["Gap"])
        if gap <= 0:
            return 0
        eff = float(row["Effective Forecasted"])
        rp = float(row["Reorder point"])
        if eff > rp and row["Units/week adj"] < 0.25:
            return 0
        qty = round_up_to_step(gap, int(s.rounding_step))
        if qty <= 0:
            return 0
        return max(int(s.moq), qty)

    df["Engine Suggested"] = df.apply(suggest, axis=1).astype(int)
    df.drop(columns=["_sold_base"], inplace=True)

    return df


# -----------------------------
# Excel Export
# -----------------------------

def build_excel(df: pd.DataFrame, col_map: Dict[str, str], s: Settings) -> bytes:
    output = io.BytesIO()
    full_df = df.copy()

    supplier_col = col_map.get("supplier_code")
    product_type_col = col_map.get("product_type")
    product_name_col = col_map.get("product_name")
    color_col = col_map.get("color_name")
    photo_col = col_map.get("photo")
    variant_col = col_map["variant_sku"]

    admin_cols = []
    if supplier_col and supplier_col in full_df.columns:
        admin_cols.append(supplier_col)

    admin_cols += ["Our Code", "Master Code", "Color Code"]

    if product_type_col and product_type_col in full_df.columns:
        admin_cols.append(product_type_col)
    if product_name_col and product_name_col in full_df.columns:
        admin_cols.append(product_name_col)
    if color_col and color_col in full_df.columns:
        admin_cols.append(color_col)

    admin_cols += [
        variant_col,
        "Size",
        "Odoo Qty Delivered",
        "Forecasted (Odoo)",
        "Effective Forecasted",
        "Units/week adj",
        "Demand units (target)",
        "Gap",
        "Reorder point",
        "Engine Suggested",
        "Override Qty",
        "Final Qty",
    ]

    if photo_col and photo_col in full_df.columns:
        admin_cols.append(photo_col)

    admin = full_df.copy()
    if "Override Qty" not in admin.columns:
        admin["Override Qty"] = np.nan
    if "Final Qty" not in admin.columns:
        admin["Final Qty"] = np.nan

    admin = admin[[c for c in admin_cols if c in admin.columns]]

    group_cols = []
    if supplier_col and supplier_col in admin.columns:
        group_cols.append(supplier_col)
    group_cols.append("Our Code")
    if color_col and color_col in admin.columns:
        group_cols.append(color_col)
    if product_name_col and product_name_col in admin.columns:
        group_cols.append(product_name_col)

    group_df = admin[group_cols].drop_duplicates()

    supplier_rows = group_df.copy()
    if supplier_col and supplier_col in supplier_rows.columns:
        supplier_rows.rename(columns={supplier_col: "Supplier Code"}, inplace=True)
    if color_col and color_col in supplier_rows.columns:
        supplier_rows.rename(columns={color_col: "Color"}, inplace=True)
    if product_name_col and product_name_col in supplier_rows.columns:
        supplier_rows.rename(columns={product_name_col: "Name"}, inplace=True)

    po_rows = group_df.copy()
    if supplier_col and supplier_col in po_rows.columns:
        po_rows.rename(columns={supplier_col: "Supplier Code"}, inplace=True)
    if color_col and color_col in po_rows.columns:
        po_rows.rename(columns={color_col: "Color"}, inplace=True)
    if product_name_col and product_name_col in po_rows.columns:
        po_rows.drop(columns=[product_name_col], inplace=True)

    po_rows["Notes"] = ""

    sizes = sorted([x for x in admin["Size"].dropna().unique()], key=lambda z: int(z) if str(z).isdigit() else str(z))
    size_headers = []
    for z in sizes:
        zs = str(z).strip()
        if zs.isdigit():
            size_headers.append(int(zs))
    size_headers = sorted(set(size_headers))

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book

        fmt_header = workbook.add_format({"bold": True, "bg_color": "#E6E6E6", "border": 1, "align": "center", "valign": "vcenter"})
        fmt_header_dark = workbook.add_format({"bold": True, "bg_color": "#111111", "font_color": "#FFFFFF", "border": 1, "align": "center", "valign": "vcenter"})
        fmt_cell = workbook.add_format({"border": 1})
        fmt_int = workbook.add_format({"border": 1, "num_format": "0"})
        fmt_float = workbook.add_format({"border": 1, "num_format": "0.00"})
        fmt_link = workbook.add_format({"border": 1, "font_color": "blue", "underline": 1})
        fmt_highlight = workbook.add_format({"border": 1, "bg_color": "#FFF2CC"})
        fmt_final = workbook.add_format({"border": 1, "bg_color": "#D9E1F2"})
        fmt_notes = workbook.add_format({"border": 1, "bg_color": "#F8F8F8"})

        settings_df = pd.DataFrame(
            [
                ["year_tag", s.year_tag],
                ["lead_time_weeks", s.lead_time_weeks],
                ["target_weeks", s.target_weeks],
                ["safety_buffer", s.safety_buffer],
                ["rounding_step", s.rounding_step],
                ["moq", s.moq],
                ["sales_window_weeks", s.sales_window_weeks],
            ],
            columns=["Setting", "Value"],
        )
        settings_df.to_excel(writer, sheet_name="SETTINGS", index=False)
        ws_settings = writer.sheets["SETTINGS"]
        ws_settings.autofilter(0, 0, len(settings_df), 1)
        ws_settings.freeze_panes(1, 0)

        full_df.to_excel(writer, sheet_name="FULL", index=False)
        ws_full = writer.sheets["FULL"]
        ws_full.autofilter(0, 0, len(full_df), len(full_df.columns) - 1)
        ws_full.freeze_panes(1, 0)

        sheet_admin = "ADMIN"
        ws_admin = workbook.add_worksheet(sheet_admin)
        writer.sheets[sheet_admin] = ws_admin

        for j, col in enumerate(admin.columns):
            header_fmt = fmt_header_dark if col in ("Engine Suggested", "Override Qty", "Final Qty") else fmt_header
            ws_admin.write(0, j, col, header_fmt)

        col_idx = {c: i for i, c in enumerate(admin.columns)}
        n_rows = len(admin)

        for i in range(n_rows):
            excel_r = i + 1
            for j, col in enumerate(admin.columns):
                val = admin.iat[i, j]

                if col == "Override Qty":
                    ws_admin.write_blank(excel_r, j, None, fmt_highlight)
                    continue

                if col == "Final Qty":
                    c_override = excel_col(col_idx["Override Qty"])
                    c_engine = excel_col(col_idx["Engine Suggested"])
                    # Greek Excel separator ';'
                    formula = f'=IF(ISNUMBER({c_override}{excel_r+1});{c_override}{excel_r+1};{c_engine}{excel_r+1})'
                    ws_admin.write_formula(excel_r, j, formula, fmt_final)
                    continue

                if col == "Engine Suggested":
                    ws_admin.write_number(excel_r, j, _safe_int(val, 0), fmt_highlight)
                elif col in ("Odoo Qty Delivered", "Forecasted (Odoo)", "Effective Forecasted", "Demand units (target)", "Gap", "Reorder point", "Units/week adj"):
                    ws_admin.write_number(excel_r, j, _safe_float(val, 0.0), fmt_float)
                elif col == "Size":
                    ws_admin.write_number(excel_r, j, _safe_int(val, 0), fmt_int)
                else:
                    if isinstance(val, str) and val.startswith("http"):
                        ws_admin.write_url(excel_r, j, val, fmt_link, string=val)
                    else:
                        ws_admin.write(excel_r, j, "" if pd.isna(val) else val, fmt_cell)

        ws_admin.autofilter(0, 0, n_rows, len(admin.columns) - 1)
        ws_admin.freeze_panes(1, 0)

        def abs_range(colname: str) -> str:
            j = col_idx[colname]
            col_letter = excel_col(j)
            return f"'{sheet_admin}'!${col_letter}$2:${col_letter}${n_rows+1}"

        rng_supplier = abs_range(supplier_col) if supplier_col and supplier_col in col_idx else None
        rng_our = abs_range("Our Code")
        rng_color = abs_range(color_col) if color_col and color_col in col_idx else None
        rng_size = abs_range("Size")
        rng_final = abs_range("Final Qty")

        # SUPPLIER
        sup_sheet = "SUPPLIER"
        ws_sup = workbook.add_worksheet(sup_sheet)
        writer.sheets[sup_sheet] = ws_sup

        sup_base_cols = []
        if "Supplier Code" in supplier_rows.columns:
            sup_base_cols.append("Supplier Code")
        sup_base_cols.append("Our Code")
        if "Name" in supplier_rows.columns:
            sup_base_cols.append("Name")
        if "Color" in supplier_rows.columns:
            sup_base_cols.append("Color")

        c = 0
        for col in sup_base_cols:
            ws_sup.write(0, c, col, fmt_header)
            c += 1
        for sz in size_headers:
            ws_sup.write(0, c, sz, fmt_header)
            c += 1
        ws_sup.write(0, c, "PHOTO", fmt_header)

        sup_col_idx = {sup_base_cols[k]: k for k in range(len(sup_base_cols))}

        for i in range(len(supplier_rows)):
            excel_r = i + 1
            c = 0
            for col in sup_base_cols:
                v = supplier_rows.iloc[i][col]
                ws_sup.write(excel_r, c, "" if pd.isna(v) else v, fmt_cell)
                c += 1

            supplier_cell = None
            if "Supplier Code" in sup_col_idx:
                supplier_cell = f"${excel_col(sup_col_idx['Supplier Code'])}${excel_r+1}"
            our_cell = f"${excel_col(sup_col_idx['Our Code'])}${excel_r+1}"
            color_cell = f"${excel_col(sup_col_idx['Color'])}${excel_r+1}" if "Color" in sup_col_idx else None

            for sz in size_headers:
                sz_header_cell = f"{excel_col(c)}$1"
                parts = [f"=SUMIFS({rng_final};{rng_size};{sz_header_cell};{rng_our};{our_cell}"]
                if supplier_cell and rng_supplier:
                    parts.append(f";{rng_supplier};{supplier_cell}")
                if color_cell and rng_color:
                    parts.append(f";{rng_color};{color_cell}")
                parts.append(")")
                ws_sup.write_formula(excel_r, c, "".join(parts), fmt_int)
                c += 1

            photo_val = ""
            if photo_col and photo_col in full_df.columns:
                mask = (full_df["Our Code"] == supplier_rows.iloc[i]["Our Code"])
                if "Supplier Code" in supplier_rows.columns and supplier_col and supplier_col in full_df.columns:
                    mask = mask & (full_df[supplier_col] == supplier_rows.iloc[i]["Supplier Code"])
                if "Color" in supplier_rows.columns and color_col and color_col in full_df.columns:
                    mask = mask & (full_df[color_col] == supplier_rows.iloc[i]["Color"])
                candidates = full_df.loc[mask, photo_col].dropna()
                if len(candidates) > 0:
                    photo_val = str(candidates.iloc[0])

            if photo_val.startswith("http"):
                ws_sup.write_url(excel_r, c, photo_val, fmt_link, string=photo_val)
            else:
                ws_sup.write(excel_r, c, photo_val, fmt_cell)

        ws_sup.autofilter(0, 0, len(supplier_rows), c)
        ws_sup.freeze_panes(1, 0)

        # PO READY
        po_sheet = "PO READY"
        ws_po = workbook.add_worksheet(po_sheet)
        writer.sheets[po_sheet] = ws_po

        po_base_cols = []
        if "Supplier Code" in po_rows.columns:
            po_base_cols.append("Supplier Code")
        po_base_cols.append("Our Code")
        if "Color" in po_rows.columns:
            po_base_cols.append("Color")
        po_base_cols.append("Notes")

        c = 0
        for col in po_base_cols:
            ws_po.write(0, c, col, fmt_header)
            c += 1
        for sz in size_headers:
            ws_po.write(0, c, sz, fmt_header)
            c += 1
        ws_po.write(0, c, "Total Units", fmt_header)
        c += 1
        ws_po.write(0, c, "PHOTO", fmt_header)

        po_col_idx_local = {po_base_cols[k]: k for k in range(len(po_base_cols))}

        for i in range(len(po_rows)):
            excel_r = i + 1
            c = 0
            for col in po_base_cols:
                v = po_rows.iloc[i][col]
                if col == "Notes":
                    ws_po.write(excel_r, c, "" if pd.isna(v) else v, fmt_notes)
                else:
                    ws_po.write(excel_r, c, "" if pd.isna(v) else v, fmt_cell)
                c += 1

            supplier_cell = f"${excel_col(po_col_idx_local['Supplier Code'])}${excel_r+1}" if "Supplier Code" in po_col_idx_local else None
            our_cell = f"${excel_col(po_col_idx_local['Our Code'])}${excel_r+1}"
            color_cell = f"${excel_col(po_col_idx_local['Color'])}${excel_r+1}" if "Color" in po_col_idx_local else None

            first_size_col = c
            for sz in size_headers:
                sz_header_cell = f"{excel_col(c)}$1"
                parts = [f"=SUMIFS({rng_final};{rng_size};{sz_header_cell};{rng_our};{our_cell}"]
                if supplier_cell and rng_supplier:
                    parts.append(f";{rng_supplier};{supplier_cell}")
                if color_cell and rng_color:
                    parts.append(f";{rng_color};{color_cell}")
                parts.append(")")
                ws_po.write_formula(excel_r, c, "".join(parts), fmt_int)
                c += 1

            total_formula = f"=SUM({excel_col(first_size_col)}{excel_r+1}:{excel_col(c-1)}{excel_r+1})"
            ws_po.write_formula(excel_r, c, total_formula, fmt_int)
            c += 1

            photo_val = ""
            if photo_col and photo_col in full_df.columns:
                mask = (full_df["Our Code"] == po_rows.iloc[i]["Our Code"])
                if "Supplier Code" in po_rows.columns and supplier_col and supplier_col in full_df.columns:
                    mask = mask & (full_df[supplier_col] == po_rows.iloc[i]["Supplier Code"])
                if "Color" in po_rows.columns and color_col and color_col in full_df.columns:
                    mask = mask & (full_df[color_col] == po_rows.iloc[i]["Color"])
                candidates = full_df.loc[mask, photo_col].dropna()
                if len(candidates) > 0:
                    photo_val = str(candidates.iloc[0])

            if photo_val.startswith("http"):
                ws_po.write_url(excel_r, c, photo_val, fmt_link, string=photo_val)
            else:
                ws_po.write(excel_r, c, photo_val, fmt_cell)

        ws_po.autofilter(0, 0, len(po_rows), c)
        ws_po.freeze_panes(1, 0)

        # Print settings for PO READY (fit to page width, repeat header row)
        ws_po.set_landscape()
        ws_po.fit_to_pages(1, 0)  # 1 page wide, unlimited tall
        ws_po.repeat_rows(0)      # repeat first row on each page
        ws_po.set_margins(0.3, 0.3, 0.4, 0.4)

    return output.getvalue()


# -----------------------------
# Streamlit UI
# -----------------------------

st.set_page_config(page_title="TFP Dynamic Restock", layout="wide")
st.title("TFP Dynamic Restock – Export Builder")

with st.sidebar:
    st.header("Settings (Minimal)")
    year_tag = st.text_input("year_tag (optional)", value="", help="Filters rows if a Season/Tag column contains this text (e.g., 2025).")
    lead_time_weeks = st.number_input("lead_time_weeks", min_value=0.0, value=6.0, step=0.5)
    target_weeks = st.number_input("target_weeks (coverage)", min_value=0.0, value=10.0, step=0.5)
    safety_buffer = st.number_input("safety_buffer (0.15 = +15%)", min_value=0.0, max_value=2.0, value=0.15, step=0.05)
    rounding_step = st.number_input("rounding_step", min_value=1, value=1, step=1)
    moq = st.number_input("moq", min_value=0, value=0, step=1)
    sales_window_weeks = st.number_input("sales_window_weeks", min_value=1.0, value=8.0, step=1.0)

    st.divider()
    st.caption("Auto column mapping. If something isn't detected, rename columns to common names (SKU/Variant SKU, Color, Photo, Forecasted, Incoming, Qty Delivered).")

settings = Settings(
    year_tag=year_tag,
    lead_time_weeks=float(lead_time_weeks),
    target_weeks=float(target_weeks),
    safety_buffer=float(safety_buffer),
    rounding_step=int(rounding_step),
    moq=int(moq),
    sales_window_weeks=float(sales_window_weeks),
)

st.subheader("1) Upload input")
uploaded = st.file_uploader("Upload your base export (CSV or XLSX)", type=["csv", "xlsx", "xls"])

if not uploaded:
    st.info("Upload a CSV/XLSX export to start.")
    st.stop()

try:
    base_df = read_table(uploaded)
    base_df = normalize_input(base_df)
    base_df, col_map = canonicalize_columns(base_df)

    with st.expander("Detected columns (auto-mapping)", expanded=False):
        st.write(col_map)

    computed = compute_engine(base_df, col_map, settings)

    st.subheader("2) Preview (ADMIN view)")
    preview_cols = [c for c in [
        col_map.get("supplier_code"),
        "Our Code",
        col_map.get("product_type"),
        col_map.get("product_name"),
        col_map.get("color_name"),
        col_map.get("variant_sku"),
        "Size",
        "Odoo Qty Delivered",
        "Forecasted (Odoo)",
        "Effective Forecasted",
        "Units/week adj",
        "Demand units (target)",
        "Gap",
        "Reorder point",
        "Engine Suggested",
    ] if c and c in computed.columns]

    st.dataframe(computed[preview_cols].head(200), use_container_width=True, height=520)

    st.subheader("3) Export")
    excel_bytes = build_excel(computed, col_map, settings)

    st.download_button(
        label="Download Excel (FULL + ADMIN + SUPPLIER + PO READY)",
        data=excel_bytes,
        file_name="tfp_dynamic_restock_export.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.markdown(
        """
**Ερώτηση: Αν αλλάξω ποσότητες στο ADMIN, ενημερώνονται SUPPLIER & PO READY;**

Ναι. Το **SUPPLIER** και το **PO READY** είναι χτισμένα με **SUMIFS formulas** πάνω στο **ADMIN → Final Qty**.
Άρα όταν συμπληρώνεις **Override Qty** στο ADMIN, αλλάζει το Final Qty και ενημερώνονται αυτόματα τα άλλα tabs.

Προσοχή μόνο σε αυτό:
- Αν **προσθέσεις/αφαιρέσεις γραμμές** χειροκίνητα στο Excel, τα SUMIFS ranges είναι “σταθερά” στις γραμμές που εξήγαγε το app.
  Σε τέτοια περίπτωση είτε επεκτείνεις ranges, είτε ξανατρέχεις export.
"""
    )

except Exception as e:
    st.error(f"Error: {e}")
