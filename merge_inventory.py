"""재고표.xls와 사원명단.xls를 '사번' 기준으로 합쳐 실사용자 이름을 붙인다."""
import pandas as pd
import xlwt

inv = pd.read_excel("재고표.xls", dtype={"사번": str})
roster = pd.read_excel("사원명단.xls", dtype={"사번": str})

# 앞뒤 공백 제거 (수작업 입력 데이터에서 매칭 실패의 흔한 원인)
inv["사번"] = inv["사번"].str.strip()
roster["사번"] = roster["사번"].str.strip()

# 재고표 기준 left merge: 재고 행은 모두 유지하고, 사원명단에 있으면 성함·부서를 붙임
merged = inv.merge(roster, on="사번", how="left", validate="many_to_one", indicator=True)

merged["비고"] = ""
merged.loc[merged["사번"].isna(), "비고"] = "미배정(재고)"
merged.loc[merged["사번"].notna() & (merged["_merge"] == "left_only"), "비고"] = "사원명단에 없는 사번"
merged = merged.drop(columns="_merge")

# 재고 데이터 바로 옆에 실사용자 이름이 오도록 컬럼 순서 고정
merged = merged[["자산번호", "사번", "성함", "부서", "비고"]]

# pandas는 .xls 쓰기를 지원하지 않으므로 xlwt로 직접 저장
out = "재고표_사용자매칭.xls"
hdr = xlwt.easyxf("font: name Arial, bold on, colour white, height 200;"
                  "pattern: pattern solid, fore_colour dark_blue; align: horiz center;"
                  "borders: left thin, right thin, top thin, bottom thin")
body = xlwt.easyxf("font: name Arial, height 200; align: horiz center;"
                   "borders: left thin, right thin, top thin, bottom thin")
wb = xlwt.Workbook(encoding="utf-8")
ws = wb.add_sheet("재고_사용자")
for j, (col, w) in enumerate(zip(merged.columns, [18, 12, 10, 12, 20])):
    ws.write(0, j, col, hdr)
    ws.col(j).width = 256 * w
for i, row in enumerate(merged.itertuples(index=False), 1):
    for j, v in enumerate(row):
        ws.write(i, j, "" if pd.isna(v) else v, body)
ws.set_panes_frozen(True)
ws.set_horz_split_pos(1)
ws.set_remove_splits(True)
wb.save(out)

print(f"총 {len(merged)}건 → {out}")
print(merged["비고"].replace("", "매칭 완료").value_counts().to_string())
print(merged.head(10).to_string(index=False))
