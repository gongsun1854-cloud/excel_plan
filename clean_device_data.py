"""장비목록 전처리
- 시리얼번호: 모든 공백 제거(탭·전각·NBSP 포함) + 대문자 변환
- MAC주소: 구분자 제거 후 AA:BB:CC:DD:EE:FF 형식으로 강제 변환, 변환 불가 시 표시
"""
import re
import pandas as pd
import xlwt

INPUT = "장비목록_원본.xls"
OUTPUT = "장비목록_정제.xls"


def clean_serial(v):
    if pd.isna(v):
        return ""
    return re.sub(r"\s+", "", str(v)).upper()    # \s는 탭·전각 공백·NBSP까지 포함


def clean_mac(v):
    """(정제값, 오류사유) 반환. 변환 불가하면 원본을 공백 제거·대문자로만 남긴다."""
    if pd.isna(v) or not str(v).strip():
        return "", "값 없음"
    raw = re.sub(r"\s+", "", str(v)).upper()

    # 구분자가 있으면 그룹 단위로 앞자리 0을 보충 (예: 0:1B:... → 00:1B:...)
    groups = re.split(r"[:\-.]", raw)
    if len(groups) == 6:
        groups = [g.zfill(2) for g in groups]
    hx = "".join(groups)

    if re.search(r"[^0-9A-F]", hx):
        return raw, "16진수가 아닌 문자 포함"
    if len(hx) != 12:
        return raw, f"자릿수 오류({len(hx)}자리)"
    return ":".join(hx[i:i + 2] for i in range(0, 12, 2)), ""


# dtype=str: '001122334455' 같은 값이 숫자로 읽혀 앞자리 0이 사라지는 것을 방지
df = pd.read_excel(INPUT, dtype=str, keep_default_na=False)
before = df[["시리얼번호", "MAC주소"]].copy()

df["시리얼번호"] = df["시리얼번호"].map(clean_serial)
df[["MAC주소", "MAC_오류"]] = df["MAC주소"].apply(lambda v: pd.Series(clean_mac(v)))

# 저장 (.xls) — MAC 변환 실패 행은 빨간 배경으로 표시
hdr = xlwt.easyxf("font: name Arial, bold on, colour white, height 200;"
                  "pattern: pattern solid, fore_colour dark_blue; align: horiz center;"
                  "borders: left thin, right thin, top thin, bottom thin")
base = "font: name Arial, height 200; borders: left thin, right thin, top thin, bottom thin"
body, bad = xlwt.easyxf(base), xlwt.easyxf(base + "; pattern: pattern solid, fore_colour rose")
wb = xlwt.Workbook(encoding="utf-8")
ws = wb.add_sheet("장비목록_정제")
for j, (c, w) in enumerate(zip(df.columns, [18, 10, 28, 10, 16, 20, 22])):
    ws.write(0, j, c, hdr)
    ws.col(j).width = 256 * w
for i, row in enumerate(df.itertuples(index=False), 1):
    st = bad if row.MAC_오류 else body
    for j, v in enumerate(row):
        ws.write(i, j, v, st)
ws.set_panes_frozen(True)
ws.set_horz_split_pos(1)
ws.set_remove_splits(True)
try:
    wb.save(OUTPUT)
except PermissionError:
    raise SystemExit(f"[저장 실패] '{OUTPUT}' 파일이 엑셀에서 열려 있습니다. 파일을 닫고 다시 실행하세요.")

changed_sn = (before["시리얼번호"] != df["시리얼번호"]).sum()
changed_mac = (before["MAC주소"] != df["MAC주소"]).sum()
errors = df[df["MAC_오류"] != ""]
print(f"총 {len(df)}건 → {OUTPUT}")
print(f"시리얼번호 수정: {changed_sn}건 / MAC주소 수정: {changed_mac}건 / MAC 변환 불가: {len(errors)}건")
if len(errors):
    print(errors[["자산번호", "MAC주소", "MAC_오류"]].to_string(index=False))
