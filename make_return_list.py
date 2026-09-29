"""테스트용 반납대상자.xls 생성 (IT자산관리_실습데이터.xlsx의 폐기예정·수리필요 기기 중 일부)"""
import pandas as pd
import xlwt

src = pd.read_excel("IT자산관리_실습데이터.xlsx").dropna(subset=["사용자"])
targets = src[src["상태"].isin(["폐기예정", "수리필요"])]

# 기기가 여러 대인 사람도 섞이도록 사람 단위로 8명 선택
people = targets["사용자"].drop_duplicates().sample(8, random_state=11)
out = (targets[targets["사용자"].isin(people)]
       .rename(columns={"사용자": "성함"})[["성함", "모델명", "자산번호"]]
       .sort_values(["성함", "자산번호"]))

hdr = xlwt.easyxf("font: name Arial, bold on, colour white, height 200;"
                  "pattern: pattern solid, fore_colour dark_blue; align: horiz center;"
                  "borders: left thin, right thin, top thin, bottom thin")
body = xlwt.easyxf("font: name Arial, height 200; align: horiz center;"
                   "borders: left thin, right thin, top thin, bottom thin")
wb = xlwt.Workbook(encoding="utf-8")
ws = wb.add_sheet("반납대상자")
for j, (col, w) in enumerate(zip(out.columns, [10, 28, 18])):
    ws.write(0, j, col, hdr)
    ws.col(j).width = 256 * w
for i, row in enumerate(out.itertuples(index=False), 1):
    for j, v in enumerate(row):
        ws.write(i, j, v, body)
wb.save("반납대상자.xls")
print(f"반납대상자.xls: {out['성함'].nunique()}명 / {len(out)}건")
print(out.to_string(index=False))
