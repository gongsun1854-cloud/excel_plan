"""테스트용 장비목록_원본.xls 생성: 시리얼번호·MAC주소를 일부러 지저분하게 만든다."""
import random
import pandas as pd
import xlwt

random.seed(99)
src = pd.read_excel("IT자산관리_실습데이터.xlsx").head(60)


def dirty_serial(s):
    s = random.choice([s, s.lower(), s.capitalize(), "".join(c.lower() if random.random() < .5 else c for c in s)])
    r = random.random()
    if r < .15:
        s = f"  {s}  "                              # 앞뒤 공백
    elif r < .30:
        s = f"{s[:4]} {s[4:8]} {s[8:]}"             # 중간 공백
    elif r < .38:
        s = s[:6] + "\t" + s[6:]                    # 탭
    elif r < .44:
        s = s[:5] + "　" + s[5:]                # 전각 공백
    elif r < .50:
        s = " " + s + " "                 # NBSP (웹에서 복사한 경우)
    return s


def dirty_mac():
    hx = "".join(random.choices("0123456789ABCDEF", k=12))
    if random.random() < .15:
        hx = "00" + hx[2:]                          # 숫자로 오인되기 쉬운 형태
    p = [hx[i:i + 2] for i in range(0, 12, 2)]
    fmt = random.choice([
        lambda: ":".join(p), lambda: ":".join(p).lower(), lambda: "-".join(p),
        lambda: "-".join(p).lower(), lambda: hx, lambda: hx.lower(),
        lambda: ".".join([hx[0:4], hx[4:8], hx[8:12]]).lower(),   # 시스코 형식
        lambda: " ".join(p), lambda: f" {':'.join(p)} ",
        lambda: ":".join(p[:3]) + "-" + "-".join(p[3:]),          # 구분자 혼용
        lambda: ":".join(x.lstrip("0") or "0" for x in p),        # 앞자리 0 누락
    ])
    return fmt()


rows = []
for r in src.itertuples(index=False):
    rows.append([r.자산번호, r.기기종류, r.모델명, r.사용자 if isinstance(r.사용자, str) else "",
                 dirty_serial(r.시리얼번호), dirty_mac()])

# 복구 불가능한 MAC 몇 건 (자릿수 부족·초과, 16진수 아닌 문자, 빈 값)
for i, bad in zip(random.sample(range(len(rows)), 4),
                  ["AA:BB:CC:DD:EE", "AA-BB-CC-DD-EE-FF-11", "ZZ:11:22:33:44:55", ""]):
    rows[i][5] = bad

hdr = xlwt.easyxf("font: name Arial, bold on, colour white, height 200;"
                  "pattern: pattern solid, fore_colour dark_blue; align: horiz center;"
                  "borders: left thin, right thin, top thin, bottom thin")
body = xlwt.easyxf("font: name Arial, height 200; borders: left thin, right thin, top thin, bottom thin")
wb = xlwt.Workbook(encoding="utf-8")
ws = wb.add_sheet("장비목록")
cols = ["자산번호", "기기종류", "모델명", "사용자", "시리얼번호", "MAC주소"]
for j, (c, w) in enumerate(zip(cols, [18, 10, 28, 10, 22, 24])):
    ws.write(0, j, c, hdr)
    ws.col(j).width = 256 * w
for i, row in enumerate(rows, 1):
    for j, v in enumerate(row):
        ws.write(i, j, v, body)
wb.save("장비목록_원본.xls")
print(f"장비목록_원본.xls: {len(rows)}건")
