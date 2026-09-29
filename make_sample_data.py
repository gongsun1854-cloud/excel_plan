"""재고표.xls / 사원명단.xls 샘플 파일 생성 (IT자산관리_실습데이터.xlsx 기반)"""
import random
import pandas as pd
import xlwt

random.seed(7)
src = pd.read_excel("IT자산관리_실습데이터.xlsx")

# 사원명단: 실습데이터의 사용자·부서 + 기기 미보유 직원 몇 명
staff = src.dropna(subset=["사용자"]).drop_duplicates("사용자")[["사용자", "부서"]]
extra = pd.DataFrame({"사용자": ["문가은", "배정훈", "허혜진", "남동현", "심보람"],
                      "부서": ["인사팀", "재무팀", "기획팀", "개발팀", "영업1팀"]})
staff = pd.concat([staff, extra], ignore_index=True).sample(frac=1, random_state=7).reset_index(drop=True)
staff["사번"] = [f"E{random.choice(range(2015, 2027))}{i:03d}" for i in range(1, len(staff) + 1)]
roster = staff.rename(columns={"사용자": "성함"})[["사번", "성함", "부서"]].sort_values("사번")

# 재고표: 자산번호 + 사번 (미배정 재고는 사번 공란, 일부는 퇴사자 사번)
name_to_id = dict(zip(staff["사용자"], staff["사번"]))
inv = pd.DataFrame({"자산번호": src["자산번호"], "사번": src["사용자"].map(name_to_id)})
retired = ["E2018901", "E2020902", "E2021903"]  # 사원명단에 없는 사번 (퇴사자)
for idx in random.sample(list(inv[inv["사번"].notna()].index), 4):
    inv.at[idx, "사번"] = random.choice(retired)


def save_xls(df, path, sheet, widths):
    hdr = xlwt.easyxf("font: name Arial, bold on, colour white, height 200;"
                      "pattern: pattern solid, fore_colour dark_blue; align: horiz center;"
                      "borders: left thin, right thin, top thin, bottom thin")
    body = xlwt.easyxf("font: name Arial, height 200; align: horiz center;"
                       "borders: left thin, right thin, top thin, bottom thin")
    wb = xlwt.Workbook(encoding="utf-8")
    ws = wb.add_sheet(sheet)
    for j, (col, w) in enumerate(zip(df.columns, widths)):
        ws.write(0, j, col, hdr)
        ws.col(j).width = 256 * w
    for i, row in enumerate(df.itertuples(index=False), 1):
        for j, v in enumerate(row):
            ws.write(i, j, "" if pd.isna(v) else v, body)
    wb.save(path)


save_xls(inv, "재고표.xls", "재고표", [18, 12])
save_xls(roster, "사원명단.xls", "사원명단", [12, 10, 12])
print(f"재고표.xls: {len(inv)}건 / 사원명단.xls: {len(roster)}명")
