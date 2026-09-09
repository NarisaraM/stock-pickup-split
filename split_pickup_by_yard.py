#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
split_pickup_by_yard.py
------------------------------------------------------------------------
อ่านไฟล์ Excel (Pickup List / BKG) ทุกไฟล์ในโฟลเดอร์ แล้ว "แยกข้อมูลตามลานลากตู้"
(คอลัมน์ Pickup / Pickup Name) ออกมาเป็นไฟล์ Excel แยกกันคนละไฟล์ต่อ 1 ลาน

แนวคิด: ดึงข้อมูลด้วยไลบรารีล้วน ๆ ไม่ใช้ AI
    - อ่าน .xls  ด้วย pandas + xlrd
    - อ่าน .xlsx ด้วย pandas + openpyxl
    - ถ้ามีไฟล์ .pdf ในโฟลเดอร์ จะใช้ pdfplumber ดึง text ออกมาเป็นไฟล์ .txt ให้ด้วย
      (ขั้นตอนนี้เป็นแค่การ "ดึงข้อความดิบ" ไม่มีการวิเคราะห์ด้วย AI)
  จากนั้นจึงค่อยเอาผลลัพธ์ที่ได้ไปให้คนหรือ AI วิเคราะห์ต่อ

ผลลัพธ์ต่อ 1 ไฟล์ต้นทาง:
    output/<ชื่อไฟล์>__000_สรุปรวมทุกลาน.xlsx     -> ตารางสรุปว่าแต่ละลานมีกี่งาน / กี่ตู้
    output/<ชื่อไฟล์>__<รหัสลาน>_<ชื่อลาน>.xlsx    -> 1 ไฟล์ต่อ 1 ลาน
        ชีต "รายการ" = รายการ booking ทั้งหมดของลานนั้น (จัดรูปแบบให้อ่านง่าย)
        ชีต "สรุป"   = นับจำนวนงาน/ตู้ แยกตามบริษัทเจ้าของสินค้า (ORG CUST)

การจัดรูปแบบให้อ่านง่าย:
    - แช่ (freeze) หัวตาราง + ใส่ AutoFilter
    - ปรับความกว้างคอลัมน์อัตโนมัติ, ตัดคำ (wrap) เฉพาะคอลัมน์หมายเหตุยาว ๆ
    - แปลง ETD (2026092301 30) -> "2026-09-23 01:30"
    - แปลง TRAN DT (20260918)   -> "2026-09-18"

วิธีใช้:
    python split_pickup_by_yard.py                       # อ่านโฟลเดอร์ที่สคริปต์อยู่
    python split_pickup_by_yard.py "D:\\path\\to\\folder"  # ระบุโฟลเดอร์
    python split_pickup_by_yard.py "D:\\folder" -o "D:\\out"

ต้องมีไลบรารี: pandas, xlrd (อ่าน .xls), openpyxl (เขียน .xlsx)  [pdfplumber = ถ้ามี]
    pip install pandas xlrd openpyxl pdfplumber
------------------------------------------------------------------------
"""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

import pandas as pd

# ----------------------------------------------------------------------
# ตั้งค่า
# ----------------------------------------------------------------------

# คอลัมน์ที่ใช้ระบุ "ลานลากตู้"
YARD_CODE_COL = "Pickup"        # รหัสลาน เช่น BKK27, LCHY5
YARD_NAME_COL = "Pickup Name"   # ชื่อลาน เช่น B.C. DEPOT CO.,LTD.

# รวมลานเหล่านี้เข้าเป็นกลุ่มเดียว (ออกไฟล์เดียว + แถวเดียวในไฟล์สรุปรวม)
# {รหัส Pickup เดิม (ตัวพิมพ์ใหญ่): (รหัสที่ใช้แสดง, ชื่อลานที่ใช้แสดง)}
# หมายเหตุ: คอลัมน์ Pickup / Pickup Name ในแต่ละแถวยังคงค่าเดิมไว้ (ไม่ถูกเขียนทับ)
YARD_MERGE = {
    "BKK01": ("BKK01+04", "PAT TERMINAL 1&2 (PORT AUTHORITY OF THAILAND)"),
    "BKK04": ("BKK01+04", "PAT TERMINAL 1&2 (PORT AUTHORITY OF THAILAND)"),
}

# ถ้าช่องเหล่านี้ "ว่างจริง" (None / NaN / ช่องว่างล้วน) ให้เติมข้อความนี้อัตโนมัติ
# — ไม่แตะช่องที่มีข้อความอยู่แล้ว
FILL_WHEN_BLANK = {
    "TRAFFIC ORDER": "GOOD AND CLEAN CONTAINER",
}

# คอลัมน์ COMMON REMARK จะถูกตัดออกจากไฟล์ต่อลาน
# ยกเว้นลาน (รหัส Pickup เดิม) เหล่านี้ ให้คงคอลัมน์ไว้ — ตำแหน่งอยู่หลัง TRAFFIC ORDER
COMMON_REMARK_KEEP = {"BKK01", "BKK02", "BKK04", "LCH55"}

# ลำดับคอลัมน์ที่อยากให้แสดงในไฟล์ผลลัพธ์ (คอลัมน์อื่นที่ไม่อยู่ในนี้จะต่อท้ายให้เอง)
PREFERRED_ORDER = [
    "BK No", "ETD", "VSL", "VOY", "POR", "LOD", "DIS", "TPSZ",
    "Pickup", "Pickup Name", "TRAN DT",
    "DOC CUST", "ORG CUST", "COMMODITY", "TRAFFIC ORDER", "COMMON REMARK",
    "GP22", "GP42", "GP45", "RE22", "RE45", "UT22", "UT42", "PC22", "PC42",
    "Pickup Qty", "Return",
]

# คอลัมน์ข้อความยาว -> ให้ตัดคำ (wrap text) และตั้งความกว้างคงที่
WRAP_COLS = ["Pickup Name", "COMMODITY", "TRAFFIC ORDER", "COMMON REMARK", "DOC CUST", "ORG CUST"]

# คอลัมน์จำนวนตู้ (เอาไว้รวมยอดในชีตสรุป)
QTY_COLS = ["GP22", "GP42", "GP45", "RE22", "RE45", "UT22", "UT42", "PC22", "PC42"]


# ----------------------------------------------------------------------
# ตัวช่วยทั่วไป
# ----------------------------------------------------------------------

def sanitize_filename(text: str, maxlen: int = 60) -> str:
    """ตัดอักขระที่ตั้งชื่อไฟล์ไม่ได้ออก"""
    text = str(text).strip()
    text = re.sub(r'[\\/:*?"<>|\r\n\t]+', " ", text)
    text = re.sub(r"\s+", " ", text).strip(" .")
    return (text[:maxlen].strip() or "NA")


def make_unique_columns(names) -> list[str]:
    """เปลี่ยนชื่อคอลัมน์ที่ซ้ำกันให้ไม่ซ้ำ แบบเดียวกับ pandas (name, name.1, name.2 ...)"""
    seen: dict[str, int] = {}
    out: list[str] = []
    for n in names:
        n = "Unnamed" if (n is None or (isinstance(n, float) and pd.isna(n))) else str(n).strip()
        if n in seen:
            seen[n] += 1
            out.append(f"{n}.{seen[n]}")
        else:
            seen[n] = 0
            out.append(n)
    return out


def digits_only(value) -> str | None:
    """เอาเฉพาะตัวเลขจากค่าเช่น '202609230130.0' -> '202609230130'"""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    s = str(value).strip()
    if not s:
        return None
    s = s.split(".")[0]                 # ตัด .0 ที่ติดมาจาก float
    s = re.sub(r"\D", "", s)            # เอาเฉพาะเลข
    return s or None


def fmt_etd(value) -> str:
    """202609230130 -> 2026-09-23 01:30 (ถ้าแปลงไม่ได้ ส่งค่าเดิมกลับเป็น string)"""
    s = digits_only(value)
    if s and len(s) == 12:
        return f"{s[0:4]}-{s[4:6]}-{s[6:8]} {s[8:10]}:{s[10:12]}"
    if s and len(s) == 8:
        return f"{s[0:4]}-{s[4:6]}-{s[6:8]}"
    return "" if value is None or (isinstance(value, float) and pd.isna(value)) else str(value).strip()


def fmt_date8(value) -> str:
    """20260918 -> 2026-09-18"""
    s = digits_only(value)
    if s and len(s) == 8:
        return f"{s[0:4]}-{s[4:6]}-{s[6:8]}"
    return "" if value is None or (isinstance(value, float) and pd.isna(value)) else str(value).strip()


# ----------------------------------------------------------------------
# อ่าน Excel
# ----------------------------------------------------------------------

def find_header_row(raw: pd.DataFrame, key: str = "BK No", scan_rows: int = 20) -> int:
    """หาแถวหัวตาราง โดยดูว่าแถวไหนมีเซลล์ที่ข้อความตรงกับ key"""
    limit = min(scan_rows, len(raw))
    for i in range(limit):
        for cell in raw.iloc[i].tolist():
            if isinstance(cell, str) and cell.strip().lower() == key.lower():
                return i
    return 0  # ไม่เจอ -> สมมติว่าแถวแรกคือหัวตาราง


def read_pickup_table(path: Path) -> pd.DataFrame:
    """อ่านไฟล์ Excel 1 ไฟล์ -> DataFrame ที่หัวตารางถูกต้องและตัดแถวว่างออกแล้ว"""
    raw = pd.read_excel(path, header=None, dtype=object)   # engine เลือกเองตามนามสกุล
    hdr = find_header_row(raw)
    header = make_unique_columns(raw.iloc[hdr].tolist())
    df = raw.iloc[hdr + 1:].copy()
    df.columns = header
    df = df.reset_index(drop=True)

    # เปลี่ยนชื่อคอลัมน์ Pickup ตัวที่ 2 (จำนวนตู้ที่ต้องไปรับ) ให้ชัดเจน
    if "Pickup.1" in df.columns:
        df = df.rename(columns={"Pickup.1": "Pickup Qty"})

    # ตัดแถวที่ไม่มีเลข booking (แถวว่าง / แถวท้ายไฟล์)
    if "BK No" in df.columns:
        bk = df["BK No"].map(lambda x: "" if x is None or (isinstance(x, float) and pd.isna(x)) else str(x).strip())
        df = df[bk.ne("") & bk.str.lower().ne("nan") & bk.str.lower().ne("none")].copy()

    # ตัดคอลัมน์ที่ว่างทั้งคอลัมน์ทิ้ง
    df = df.dropna(axis=1, how="all")
    return df.reset_index(drop=True)


def _is_blank(x) -> bool:
    """True เมื่อค่าว่างจริง: None / NaN / ช่องว่างล้วน"""
    if x is None:
        return True
    if isinstance(x, float) and pd.isna(x):
        return True
    return isinstance(x, str) and x.strip() == ""


def clean_frame(df: pd.DataFrame) -> pd.DataFrame:
    """จัดรูปแบบค่าต่าง ๆ ให้อ่านง่าย"""
    df = df.copy()

    # ตัดช่องว่างหน้า-หลังของทุกคอลัมน์ข้อความ
    for col in df.columns:
        if df[col].dtype == object:
            df[col] = df[col].map(lambda x: x.strip() if isinstance(x, str) else x)

    # เติมค่าให้ช่องที่ "ว่างจริง" เท่านั้น (ไม่แตะช่องที่มีข้อความอยู่แล้ว)
    for col, value in FILL_WHEN_BLANK.items():
        if col in df.columns:
            df[col] = df[col].map(lambda x: value if _is_blank(x) else x)

    # ORG CUST ว่าง -> เติมจาก DOC CUST (แถวเดียวกัน) แล้วตัด DOC CUST ทิ้ง
    if "ORG CUST" in df.columns and "DOC CUST" in df.columns:
        blank = df["ORG CUST"].map(_is_blank)
        df.loc[blank, "ORG CUST"] = df.loc[blank, "DOC CUST"]
    if "DOC CUST" in df.columns:
        df = df.drop(columns=["DOC CUST"])

    if "ETD" in df.columns:
        df["ETD"] = df["ETD"].map(fmt_etd)
    if "TRAN DT" in df.columns:
        df["TRAN DT"] = df["TRAN DT"].map(fmt_date8)

    # คอลัมน์จำนวนตู้ -> ตัวเลขจำนวนเต็ม
    for col in QTY_COLS + ["Pickup Qty", "Return"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).astype(int)

    return df


def order_columns(df: pd.DataFrame) -> pd.DataFrame:
    """จัดลำดับคอลัมน์ตาม PREFERRED_ORDER แล้วต่อท้ายด้วยคอลัมน์ที่เหลือ"""
    front = [c for c in PREFERRED_ORDER if c in df.columns]
    rest = [c for c in df.columns if c not in front]
    return df[front + rest]


def yard_group_keys(df: pd.DataFrame) -> tuple[pd.Series, pd.Series]:
    """คืน (รหัสลาน, ชื่อลาน) สำหรับใช้ 'จัดกลุ่ม' โดยรวมลานตาม YARD_MERGE แล้ว
    (ใช้ร่วมกันทั้งไฟล์สรุปรวมและการแยกไฟล์ต่อลาน เพื่อให้ผลตรงกัน)"""
    raw_code = (df.get(YARD_CODE_COL, pd.Series(["NA"] * len(df), index=df.index))
                  .fillna("NA").astype(str).str.strip().replace("", "NA"))
    raw_name = (df.get(YARD_NAME_COL, pd.Series([""] * len(df), index=df.index))
                  .fillna("").astype(str).str.strip())

    gcode, gname = [], []
    for code, name in zip(raw_code, raw_name):
        merged = YARD_MERGE.get(code.upper())
        if merged:
            gcode.append(merged[0])
            gname.append(merged[1])
        else:
            gcode.append(code)
            gname.append(name)
    return pd.Series(gcode, index=df.index), pd.Series(gname, index=df.index)


# ----------------------------------------------------------------------
# เขียน Excel (จัดรูปแบบให้อ่านง่าย)
# ----------------------------------------------------------------------

# --- ตั้งค่ารูปแบบ (ใช้ร่วมกันทุกไฟล์/ทุกชีต) ---
FONT_NAME = "Calibri"
FONT_SIZE = 11
ROW_HEIGHT = 13

# ไฮไลท์ทั้งแถวตามค่าในคอลัมน์ TPSZ
TPSZ_YELLOW = ("20RF", "R40H")                       # ตู้ห้องเย็น
TPSZ_GREEN = ("20OT", "20FR", "40OT", "40FR")        # ตู้ open-top / flat-rack
# คำใน TRAFFIC ORDER ที่ทำให้ตัวอักษรเป็นสีแดงตัวหนา
PRECOOL_RE = re.compile(r"pre[-\s]?cool", re.IGNORECASE)


def _style_sheet(ws, df: pd.DataFrame, wrap_cols: list[str]) -> None:
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(name=FONT_NAME, size=FONT_SIZE, bold=True, color="FFFFFF")
    base_font = Font(name=FONT_NAME, size=FONT_SIZE)
    total_font = Font(name=FONT_NAME, size=FONT_SIZE, bold=True)
    precool_font = Font(name=FONT_NAME, size=FONT_SIZE, bold=True, color="FF0000")

    yellow_fill = PatternFill("solid", fgColor="FFFF00")
    green_fill = PatternFill("solid", fgColor="92D050")

    cols = list(df.columns)
    ncol = len(cols)
    col_i = {c: i + 1 for i, c in enumerate(cols)}      # ชื่อคอลัมน์ -> เลขคอลัมน์ (1-based)
    tpsz_i = col_i.get("TPSZ")
    traffic_i = col_i.get("TRAFFIC ORDER")

    # ---- หัวตาราง ----
    for ci in range(1, ncol + 1):
        cell = ws.cell(row=1, column=ci)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(vertical="center", wrap_text=True)

    ws.freeze_panes = "A2"
    if ws.max_row >= 1 and ws.max_column >= 1:
        ws.auto_filter.ref = ws.dimensions

    # ---- ความกว้างคอลัมน์ ----
    wrap_set = {c for c in wrap_cols if c in cols}
    wrap_idx = {col_i[c] for c in wrap_set}
    for ci, col in enumerate(cols, start=1):
        letter = get_column_letter(ci)
        if col in wrap_set:
            ws.column_dimensions[letter].width = 46
        else:
            sample = [str(x) for x in df[col].head(300).tolist()]
            width = max([len(str(col))] + [len(s) for s in sample]) + 2
            ws.column_dimensions[letter].width = min(max(width, 10), 34)

    # ---- แถวข้อมูล : ฟอนต์ / ไฮไลท์ / จัดข้อความ ----
    for ri in range(2, ws.max_row + 1):
        rec = df.iloc[ri - 2]
        is_total = str(rec.iloc[0]).strip() == "รวมทั้งหมด"

        row_fill = None
        if tpsz_i is not None:
            tp = str(rec.iloc[tpsz_i - 1]).upper().replace(" ", "")
            if any(k in tp for k in TPSZ_YELLOW):
                row_fill = yellow_fill
            elif any(k in tp for k in TPSZ_GREEN):
                row_fill = green_fill

        is_precool = (
            traffic_i is not None
            and bool(PRECOOL_RE.search(str(rec.iloc[traffic_i - 1])))
        )

        for ci in range(1, ncol + 1):
            cell = ws.cell(row=ri, column=ci)
            cell.font = total_font if is_total else base_font
            if row_fill is not None:
                cell.fill = row_fill
            if ci in wrap_idx:
                cell.alignment = Alignment(wrap_text=True, vertical="top")

        if is_precool:
            ws.cell(row=ri, column=traffic_i).font = precool_font

    # ---- ความสูงทุกแถว (รวมหัวตาราง) = ROW_HEIGHT ----
    for ri in range(1, ws.max_row + 1):
        ws.row_dimensions[ri].height = ROW_HEIGHT


def build_customer_summary(df: pd.DataFrame) -> pd.DataFrame:
    """สรุปจำนวนงาน/ตู้ แยกตามบริษัทเจ้าของสินค้า (ORG CUST)"""
    group_col = "ORG CUST" if "ORG CUST" in df.columns else ("DOC CUST" if "DOC CUST" in df.columns else None)
    if group_col is None:
        return pd.DataFrame()

    qty_cols = [c for c in QTY_COLS if c in df.columns]
    agg = {"BK No": "count"} if "BK No" in df.columns else {}
    for c in qty_cols:
        agg[c] = "sum"

    summary = (
        df.assign(**{group_col: df[group_col].fillna("(ไม่ระบุ)").replace("", "(ไม่ระบุ)")})
          .groupby(group_col, dropna=False)
          .agg(agg)
          .reset_index()
    )
    if "BK No" in summary.columns:
        summary = summary.rename(columns={"BK No": "จำนวนงาน"})
    if qty_cols:
        summary["รวมตู้"] = summary[qty_cols].sum(axis=1)
    sort_col = "จำนวนงาน" if "จำนวนงาน" in summary.columns else group_col
    return summary.sort_values(sort_col, ascending=False).reset_index(drop=True)


def write_yard_file(out_path: Path, df_yard: pd.DataFrame) -> None:
    df_list = order_columns(clean_frame(df_yard))

    # ตัด COMMON REMARK ออก ยกเว้นลานใน COMMON_REMARK_KEEP (คงไว้หลัง TRAFFIC ORDER)
    if "COMMON REMARK" in df_list.columns:
        codes = set(
            df_yard.get(YARD_CODE_COL, pd.Series(dtype="object"))
                   .dropna().astype(str).str.strip().str.upper()
        )
        if not (codes & COMMON_REMARK_KEEP):
            df_list = df_list.drop(columns=["COMMON REMARK"])

    df_sum = build_customer_summary(df_list)

    with pd.ExcelWriter(out_path, engine="openpyxl") as xw:
        df_list.to_excel(xw, sheet_name="รายการ", index=False)
        _style_sheet(xw.sheets["รายการ"], df_list, WRAP_COLS)
        if not df_sum.empty:
            df_sum.to_excel(xw, sheet_name="สรุป", index=False)
            _style_sheet(xw.sheets["สรุป"], df_sum, ["ORG CUST", "DOC CUST"])


def write_overview_file(out_path: Path, df_all: pd.DataFrame) -> pd.DataFrame:
    """สร้างไฟล์สรุปรวม: แต่ละลานมีกี่งาน / กี่ตู้"""
    df = clean_frame(df_all)
    gcode, gname = yard_group_keys(df)
    df = df.assign(_code=gcode, _name=gname)

    qty_cols = [c for c in QTY_COLS if c in df.columns]
    rows = []
    for c_val, g in df.groupby("_code"):
        row = {
            "รหัสลาน": c_val,
            "ชื่อลาน": g["_name"].mode().iat[0] if not g["_name"].mode().empty else "",
            "จำนวนงาน": len(g),
        }
        for c in qty_cols:
            row[c] = int(g[c].sum())
        if qty_cols:
            row["รวมตู้"] = int(g[qty_cols].sum().sum())
        if "Pickup Qty" in g.columns:
            row["Pickup Qty"] = int(g["Pickup Qty"].sum())
        if "Return" in g.columns:
            row["Return"] = int(g["Return"].sum())
        rows.append(row)

    overview = pd.DataFrame(rows).sort_values("จำนวนงาน", ascending=False).reset_index(drop=True)
    total = {"รหัสลาน": "รวมทั้งหมด", "ชื่อลาน": "", "จำนวนงาน": overview["จำนวนงาน"].sum()}
    for c in overview.columns:
        if c not in total:
            total[c] = overview[c].sum()
    overview = pd.concat([overview, pd.DataFrame([total])], ignore_index=True)
    overview["ชื่อลาน"] = overview["ชื่อลาน"].fillna("")

    with pd.ExcelWriter(out_path, engine="openpyxl") as xw:
        overview.to_excel(xw, sheet_name="สรุปรวมทุกลาน", index=False)
        _style_sheet(xw.sheets["สรุปรวมทุกลาน"], overview, ["ชื่อลาน"])
    return overview


# ----------------------------------------------------------------------
# pdfplumber (ถ้ามีไฟล์ PDF ในโฟลเดอร์)
# ----------------------------------------------------------------------

def dump_pdf_text(folder: Path, out_dir: Path) -> None:
    pdfs = sorted(folder.glob("*.pdf"))
    if not pdfs:
        return
    try:
        import pdfplumber
    except ImportError:
        print(f"  ! พบไฟล์ PDF {len(pdfs)} ไฟล์ แต่ยังไม่ได้ติดตั้ง pdfplumber "
              f"(ข้าม) -> pip install pdfplumber")
        return

    for pdf_path in pdfs:
        txt_path = out_dir / f"{pdf_path.stem}__text.txt"
        parts: list[str] = []
        with pdfplumber.open(pdf_path) as pdf:
            for i, page in enumerate(pdf.pages, start=1):
                parts.append(f"\n===== หน้า {i} / {pdf_path.name} =====\n")
                parts.append(page.extract_text() or "")
        txt_path.write_text("\n".join(parts), encoding="utf-8")
        print(f"  - ดึงข้อความจาก PDF: {pdf_path.name} -> {txt_path.name}")


# ----------------------------------------------------------------------
# main
# ----------------------------------------------------------------------

def process_excel_file(path: Path, out_dir: Path) -> None:
    print(f"\n[ไฟล์] {path.name}")
    df = read_pickup_table(path)
    if df.empty:
        print("  ! ไม่พบข้อมูลในไฟล์นี้ (ข้าม)")
        return
    if YARD_CODE_COL not in df.columns and YARD_NAME_COL not in df.columns:
        print(f"  ! ไม่พบคอลัมน์ '{YARD_CODE_COL}' หรือ '{YARD_NAME_COL}' (ข้าม)")
        return

    print(f"  อ่านได้ {len(df)} แถว, {len(df.columns)} คอลัมน์")
    stem = sanitize_filename(path.stem)

    # ไฟล์สรุปรวม (ขึ้นต้น 000 เพื่อให้เรียงอยู่บนสุด)
    overview_path = out_dir / f"{stem}__000_สรุปรวมทุกลาน.xlsx"
    overview = write_overview_file(overview_path, df)
    print(f"  + {overview_path.name}")

    # แยกไฟล์ตามลาน (รวมลานตาม YARD_MERGE)
    gcode, gname = yard_group_keys(df)

    for code, g in df.assign(_code=gcode, _name=gname).groupby("_code"):
        yard_name = g["_name"].mode().iat[0] if not g["_name"].mode().empty else ""
        label = sanitize_filename(f"{code}_{yard_name}" if yard_name else code)
        yard_path = out_dir / f"{stem}__{label}.xlsx"
        write_yard_file(yard_path, g.drop(columns=["_code", "_name"]))
        print(f"  + {yard_path.name}  ({len(g)} งาน)")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="แยกข้อมูล Pickup List ตามลานลากตู้ ออกเป็นไฟล์ Excel คนละไฟล์"
    )
    parser.add_argument(
        "folder", nargs="?", default=None,
        help="โฟลเดอร์ที่มีไฟล์ Excel (ค่าเริ่มต้น = โฟลเดอร์ที่สคริปต์นี้อยู่)",
    )
    parser.add_argument(
        "-o", "--output", default=None,
        help="โฟลเดอร์ผลลัพธ์ (ค่าเริ่มต้น = <folder>/output)",
    )
    args = parser.parse_args(argv)

    folder = Path(args.folder).expanduser().resolve() if args.folder else Path(__file__).parent.resolve()
    if not folder.is_dir():
        print(f"ไม่พบโฟลเดอร์: {folder}")
        return 1

    out_dir = Path(args.output).expanduser().resolve() if args.output else (folder / "output")
    out_dir.mkdir(parents=True, exist_ok=True)

    # หาไฟล์ Excel (ไม่รวมไฟล์ชั่วคราว ~$ และไฟล์ผลลัพธ์ของเราเอง)
    excel_files = []
    for pattern in ("*.xls", "*.xlsx", "*.xlsm"):
        for p in folder.glob(pattern):
            if p.name.startswith("~$"):
                continue
            if out_dir in p.parents or p.parent == out_dir:
                continue
            excel_files.append(p)
    excel_files = sorted(set(excel_files))

    print(f"โฟลเดอร์ต้นทาง : {folder}")
    print(f"โฟลเดอร์ผลลัพธ์: {out_dir}")
    print(f"พบไฟล์ Excel   : {len(excel_files)} ไฟล์")

    if not excel_files:
        print("ไม่พบไฟล์ Excel ในโฟลเดอร์นี้")
    for path in excel_files:
        try:
            process_excel_file(path, out_dir)
        except Exception as exc:  # noqa: BLE001
            print(f"  ! ผิดพลาดกับไฟล์ {path.name}: {exc}")

    dump_pdf_text(folder, out_dir)

    print("\nเสร็จแล้ว")
    return 0


if __name__ == "__main__":
    sys.exit(main())
