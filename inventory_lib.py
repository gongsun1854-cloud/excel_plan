"""재고 자동화 함수 모음 (1단계: 기존 스크립트 로직을 재사용 가능한 함수로 정리)

- merge_branch_files : 지점별 파일 여러 개를 하나로 합치기 (pd.concat)
- extract_repair     : 수리 대상 추출
- make_repair_pdf    : 수리 대상 목록 PDF 생성 (fpdf2)
- parse_rules / clean_serial / clean_mac / clean_data : 데이터 청소
- to_excel_bytes     : 결과를 엑셀(.xlsx) 바이트로 변환 (다운로드용)
"""
import io
import logging
import os
import re
from datetime import date

import pandas as pd
from fpdf import FPDF
from fpdf.fonts import FontFace

logging.getLogger("fontTools.subset").setLevel(logging.ERROR)  # 한글 폰트 서브셋 경고 숨김

FONT_REGULAR = r"C:\Windows\Fonts\malgun.ttf"
FONT_BOLD = r"C:\Windows\Fonts\malgunbd.ttf"
STATUS_REPAIR = "수리필요"
STATUS_DISPOSE = "폐기예정"


# ---------------------------------------------------------------- 1. 합치기
def branch_name(filename):
    """'강남_재고.xls' → '강남'"""
    stem = os.path.splitext(os.path.basename(filename))[0]
    return stem.split("_")[0]


def read_excel_file(file):
    """경로 또는 업로드 파일 객체를 읽는다. 시리얼·MAC은 앞자리 0 보존을 위해 문자로 읽음."""
    return pd.read_excel(file, dtype={"시리얼번호": str, "MAC주소": str, "사번": str})


def merge_branch_files(files):
    """여러 지점 파일을 합쳐 (합친 DataFrame, 건너뛴 파일 목록)을 반환.
    files: [(파일명, 파일객체 또는 경로), ...]
    """
    frames, skipped = [], []
    for name, f in files:
        try:
            df = read_excel_file(f)
        except Exception as e:
            skipped.append((name, f"읽기 실패: {e}"))
            continue
        if df.empty:
            skipped.append((name, "데이터 없음"))
            continue
        df.insert(0, "지점", branch_name(name))
        frames.append(df)
    if not frames:
        return pd.DataFrame(), skipped
    merged = pd.concat(frames, ignore_index=True)
    if "구매일자" in merged.columns:
        merged["구매일자"] = pd.to_datetime(merged["구매일자"], errors="coerce")
    return merged, skipped


# ---------------------------------------------------------------- 2. 추출
def extract_repair(df, min_days=None, today=None):
    """상태가 '수리필요'인 행. min_days를 주면 구매 후 그 일수 이상 지난 것만."""
    if "상태" not in df.columns:
        raise KeyError("'상태' 컬럼이 없습니다.")
    mask = df["상태"].astype(str).str.strip() == STATUS_REPAIR
    if min_days is not None and "구매일자" in df.columns:
        today = pd.Timestamp(today or date.today())
        mask &= (today - pd.to_datetime(df["구매일자"], errors="coerce")).dt.days >= min_days
    return df[mask].copy()


class _KoreanPDF(FPDF):
    def header(self):
        self.set_font("Malgun", "B", 16)
        self.cell(0, 12, "수리 대상 기기 목록", align="C", new_x="LMARGIN", new_y="NEXT")
        self.set_font("Malgun", "", 9)
        self.cell(0, 6, f"출력일: {date.today():%Y-%m-%d}", align="R", new_x="LMARGIN", new_y="NEXT")
        self.ln(2)

    def footer(self):
        self.set_y(-12)
        self.set_font("Malgun", "", 8)
        self.cell(0, 8, f"{self.page_no()} / {{nb}}", align="C")


def make_repair_pdf(df):
    """수리 대상 DataFrame을 표 형태 PDF(bytes)로 만든다."""
    cols = [c for c in ["지점", "자산번호", "기기종류", "모델명", "사용자", "부서", "구매일자"] if c in df.columns]
    widths = {"지점": 14, "자산번호": 36, "기기종류": 16, "모델명": 50, "사용자": 18, "부서": 22, "구매일자": 24}

    pdf = _KoreanPDF(orientation="P", unit="mm", format="A4")
    pdf.add_font("Malgun", "", FONT_REGULAR)
    pdf.add_font("Malgun", "B", FONT_BOLD)
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()

    pdf.set_font("Malgun", "", 10)
    pdf.cell(0, 7, f"총 {len(df)}건 · 상태 '{STATUS_REPAIR}'", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(2)

    pdf.set_font("Malgun", "", 8)
    with pdf.table(col_widths=[widths.get(c, 20) for c in cols], text_align="CENTER",
                   line_height=6,
                   headings_style=FontFace(emphasis="BOLD", color=255, fill_color=(48, 84, 150))) as table:
        table.row(cols)
        for rec in df[cols].itertuples(index=False):
            vals = []
            for c, v in zip(cols, rec):
                if pd.isna(v):
                    vals.append("")
                elif c == "구매일자":
                    vals.append(pd.Timestamp(v).strftime("%Y-%m-%d"))
                else:
                    vals.append(str(v))
            table.row(vals)

    pdf.ln(8)
    pdf.set_font("Malgun", "", 10)
    pdf.cell(0, 7, "위 기기에 대해 수리를 요청합니다.", align="C")
    return bytes(pdf.output())


# ---------------------------------------------------------------- 3. 청소
def parse_rules(text):
    """'O=0, I=1, -=' → [('O','0'), ('I','1'), ('-','')]"""
    rules = []
    for part in (text or "").split(","):
        if "=" in part:
            old, new = part.split("=", 1)
            old = old.strip()
            if old:
                rules.append((old, new.strip()))
    return rules


def clean_serial(value, upper=True, remove_space=True, rules=()):
    if pd.isna(value):
        return ""
    s = str(value)
    if remove_space:
        s = re.sub(r"\s+", "", s)          # 탭·전각 공백·NBSP 포함
    if upper:
        s = s.upper()
    for old, new in rules:
        s = s.replace(old.upper() if upper else old, new)
    return s


def clean_mac(value):
    """(정제값, 오류사유). 12자리 16진수면 AA:BB:CC:DD:EE:FF로 강제 변환."""
    if pd.isna(value) or not str(value).strip():
        return "", "값 없음"
    raw = re.sub(r"\s+", "", str(value)).upper()
    groups = re.split(r"[:\-.]", raw)
    if len(groups) == 6:
        groups = [g.zfill(2) for g in groups]
    hx = "".join(groups)
    if re.search(r"[^0-9A-F]", hx):
        return raw, "16진수가 아닌 문자 포함"
    if len(hx) != 12:
        return raw, f"자릿수 오류({len(hx)}자리)"
    return ":".join(hx[i:i + 2] for i in range(0, 12, 2)), ""


def clean_data(df, upper=True, remove_space=True, rules=(), fix_mac=True):
    """(청소된 DataFrame, 시리얼 변경 건수, MAC 변경 건수)"""
    out = df.copy()
    sn_changed = mac_changed = 0
    if "시리얼번호" in out.columns:
        before = out["시리얼번호"].fillna("").astype(str)
        out["시리얼번호"] = out["시리얼번호"].map(lambda v: clean_serial(v, upper, remove_space, rules))
        sn_changed = int((before != out["시리얼번호"]).sum())
    if fix_mac and "MAC주소" in out.columns:
        # MAC주소 컬럼이 없던 파일에서 온 행(NaN)은 검사 대상에서 제외
        has_mac = out["MAC주소"].notna()
        before = out["MAC주소"].fillna("").astype(str)
        res = out.loc[has_mac, "MAC주소"].map(clean_mac)
        out["MAC_오류"] = ""
        out.loc[has_mac, "MAC주소"] = res.str[0]
        out.loc[has_mac, "MAC_오류"] = res.str[1]
        out["MAC주소"] = out["MAC주소"].fillna("")
        mac_changed = int((before != out["MAC주소"]).sum())
    return out, sn_changed, mac_changed


# ---------------------------------------------------------------- 다운로드
def to_excel_bytes(sheets):
    """{시트명: DataFrame} → .xlsx 바이트"""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl", date_format="YYYY-MM-DD",
                        datetime_format="YYYY-MM-DD") as xw:
        for name, df in sheets.items():
            df.to_excel(xw, sheet_name=name[:31], index=False)
            ws = xw.sheets[name[:31]]
            for i, col in enumerate(df.columns, 1):
                width = max([len(str(col))] + [len(str(v)) for v in df[col].head(200)]) + 4
                ws.column_dimensions[ws.cell(1, i).column_letter].width = min(width, 40)
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
    return buf.getvalue()
