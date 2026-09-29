"""반납대상자.xls를 읽어 사람마다 '이름_반납확인서.pdf'를 생성한다."""
import os
import re
from datetime import date

import pandas as pd
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

INPUT = "반납대상자.xls"
OUT_DIR = "반납확인서"
FILE_FMT = "{name}_반납확인서.pdf"
CONFIRM_TEXT = "위 기기를 정상 반납했음을 확인합니다."

# 한글 출력용 맑은 고딕 (Windows 기본 글꼴)
pdfmetrics.registerFont(TTFont("Malgun", r"C:\Windows\Fonts\malgun.ttf"))
pdfmetrics.registerFont(TTFont("MalgunBd", r"C:\Windows\Fonts\malgunbd.ttf"))

title_st = ParagraphStyle("title", fontName="MalgunBd", fontSize=22, alignment=TA_CENTER, leading=30)
body_st = ParagraphStyle("body", fontName="Malgun", fontSize=11, leading=18)
confirm_st = ParagraphStyle("confirm", fontName="MalgunBd", fontSize=13, alignment=TA_CENTER, leading=20)
center_st = ParagraphStyle("center", fontName="Malgun", fontSize=12, alignment=TA_CENTER, leading=18)


def safe_filename(s):
    return re.sub(r'[\\/:*?"<>|]', "_", str(s).strip())


def build_pdf(name, items, path, today):
    doc = SimpleDocTemplate(path, pagesize=A4, leftMargin=25 * mm, rightMargin=25 * mm,
                            topMargin=30 * mm, bottomMargin=25 * mm,
                            title=f"{name} 반납확인서", author="전산팀")
    story = [Paragraph("전산기기 반납확인서", title_st), Spacer(1, 12 * mm)]

    info = Table([["성    함", name], ["반납일자", today.strftime("%Y년 %m월 %d일")],
                  ["반납수량", f"{len(items)}대"]], colWidths=[35 * mm, 125 * mm])
    info.setStyle(TableStyle([
        ("FONT", (0, 0), (-1, -1), "Malgun", 11),
        ("FONT", (0, 0), (0, -1), "MalgunBd", 11),
        ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#E8EEF7")),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#7F7F7F")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (0, 0), (0, -1), "CENTER"),
        ("TOPPADDING", (0, 0), (-1, -1), 6), ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
    ]))
    story += [info, Spacer(1, 8 * mm), Paragraph("■ 반납 기기 내역", body_st), Spacer(1, 2 * mm)]

    rows = [["No", "자산번호", "모델명"]]
    rows += [[i, r["자산번호"], r["모델명"]] for i, r in enumerate(items.to_dict("records"), 1)]
    dev = Table(rows, colWidths=[15 * mm, 50 * mm, 95 * mm], repeatRows=1)
    dev.setStyle(TableStyle([
        ("FONT", (0, 0), (-1, -1), "Malgun", 10),
        ("FONT", (0, 0), (-1, 0), "MalgunBd", 10),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#305496")),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#7F7F7F")),
        ("ALIGN", (0, 0), (1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 5), ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))
    story += [dev, Spacer(1, 18 * mm),
              Paragraph(CONFIRM_TEXT, confirm_st), Spacer(1, 12 * mm),
              Paragraph(today.strftime("%Y년 %m월 %d일"), center_st), Spacer(1, 14 * mm)]

    sign = Table([["반 납 자 :", name, "(서명 또는 인)"],
                  ["확 인 자 :", "", "(서명 또는 인)"]],
                 colWidths=[30 * mm, 40 * mm, 35 * mm], rowHeights=12 * mm, hAlign="RIGHT")
    sign.setStyle(TableStyle([
        ("FONT", (0, 0), (-1, -1), "Malgun", 11),
        ("TEXTCOLOR", (2, 0), (2, -1), colors.HexColor("#7F7F7F")),
        ("LINEBELOW", (1, 0), (1, -1), 0.5, colors.black),
        ("ALIGN", (1, 0), (1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "BOTTOM"),
    ]))
    story.append(sign)
    doc.build(story)


def main():
    df = pd.read_excel(INPUT, dtype=str)
    df = df.apply(lambda s: s.str.strip())
    df = df.dropna(subset=["성함"])
    os.makedirs(OUT_DIR, exist_ok=True)
    today = date.today()

    # 같은 사람의 기기는 한 장의 확인서에 모아서 출력
    for name, items in df.groupby("성함", sort=True):
        path = os.path.join(OUT_DIR, FILE_FMT.format(name=safe_filename(name)))
        build_pdf(name, items, path, today)
        print(f"생성: {path} ({len(items)}대)")
    print(f"총 {df['성함'].nunique()}명, {len(df)}건 → {OUT_DIR}\\")


if __name__ == "__main__":
    main()
