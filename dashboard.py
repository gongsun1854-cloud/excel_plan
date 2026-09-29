"""나만의 재고 웹 대시보드 (실행: streamlit run dashboard.py)

2단계: inventory_lib 함수들을 Streamlit UI와 연결
3단계: 결과물을 엑셀/PDF로 다운로드
"""
from datetime import date

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

import inventory_lib as lib

st.set_page_config(page_title="재고 대시보드", page_icon="📦", layout="wide")
st.title("📦 나만의 재고 웹 대시보드")

TODAY = date.today()


@st.cache_data(show_spinner="파일을 합치는 중...")
def load(files):
    # files: ((이름, bytes), ...) — 해시 가능한 형태로 받아 캐시
    import io
    return lib.merge_branch_files([(n, io.BytesIO(b)) for n, b in files])


# ---------------------------------------------------------------- 사이드바: 업로드
with st.sidebar:
    st.header("📁 지점별 파일 업로드")
    uploads = st.file_uploader("여러 파일을 한꺼번에 선택하세요 (예: 강남_재고.xls)",
                               type=["xlsx", "xls"], accept_multiple_files=True)
    st.caption("파일명의 '_' 앞부분이 지점명이 됩니다.")

if not uploads:
    st.info("왼쪽에서 지점별 엑셀 파일을 올리면 한 번에 합쳐서 보여줍니다.")
    st.stop()

files_key = tuple(sorted((f.name, f.getvalue()) for f in uploads))
df, skipped = load(files_key)

# 업로드 파일이 바뀌면 이전 추출·청소 결과 초기화
sig = tuple(n for n, _ in files_key)
if st.session_state.get("sig") != sig:
    st.session_state.update(sig=sig, repair=None, cleaned=None)

for name, reason in skipped:
    st.warning(f"'{name}' 건너뜀 — {reason}")
if df.empty:
    st.error("읽을 수 있는 데이터가 없습니다.")
    st.stop()

status = df["상태"].astype(str).str.strip() if "상태" in df.columns else pd.Series(dtype=str)
c1, c2, c3, c4 = st.columns(4)
c1.metric("지점 수", f"{df['지점'].nunique()}곳")
c2.metric("전체 기기 수", f"{len(df):,}대")
c3.metric("수리 필요 기기 수", f"{(status == lib.STATUS_REPAIR).sum():,}대")
c4.metric("폐기 예정 기기 수", f"{(status == lib.STATUS_DISPOSE).sum():,}대")

tab1, tab2, tab3 = st.tabs(["📊 통합 현황", "🔧 수리 대상 추출", "🧹 데이터 청소"])

# ---------------------------------------------------------------- 탭1: 합친 결과
with tab1:
    st.subheader("지점별 상태 현황")
    if "상태" in df.columns:
        pivot = pd.crosstab(df["지점"], df["상태"], margins=True, margins_name="합계")
        st.dataframe(pivot, width="stretch")

    st.subheader("통합 데이터 미리보기")
    st.caption(f"{len(uploads)}개 파일 · {len(df):,}행 × {len(df.columns)}열")
    st.dataframe(df, width="stretch", hide_index=True,
                 column_config={"구매일자": st.column_config.DateColumn(format="YYYY-MM-DD")})

    st.download_button("📥 통합 파일 엑셀 다운로드",
                       data=lib.to_excel_bytes({"통합재고": df}),
                       file_name=f"전지점_통합재고_{TODAY:%Y%m%d}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ---------------------------------------------------------------- 탭2: 수리 대상 추출 + PDF
with tab2:
    only_old = st.checkbox("구매 후 3년(1095일) 이상 지난 기기만", value=False)
    if st.button("🔧 수리 대상자 추출", type="primary"):
        try:
            st.session_state.repair = lib.extract_repair(df, min_days=1095 if only_old else None)
        except KeyError as e:
            st.error(str(e))

    repair = st.session_state.get("repair")
    if repair is not None:
        if repair.empty:
            st.info("조건에 맞는 수리 대상이 없습니다.")
        else:
            st.success(f"수리 대상 {len(repair)}건을 찾았습니다.")
            st.dataframe(repair, width="stretch", hide_index=True,
                         column_config={"구매일자": st.column_config.DateColumn(format="YYYY-MM-DD")})
            d1, d2 = st.columns(2)
            d1.download_button("📄 PDF 다운로드", data=lib.make_repair_pdf(repair),
                               file_name=f"수리대상_{TODAY:%Y%m%d}.pdf", mime="application/pdf")
            d2.download_button("📥 엑셀 다운로드", data=lib.to_excel_bytes({"수리대상": repair}),
                               file_name=f"수리대상_{TODAY:%Y%m%d}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ---------------------------------------------------------------- 탭3: 데이터 청소
with tab3:
    st.write("시리얼번호를 아래 규칙대로 수정합니다. 공백 제거 → 대문자 변환 → 치환 규칙 순서로 적용됩니다.")
    rule_text = st.text_input("변경할 시리얼 번호 규칙 (찾을값=바꿀값, 쉼표로 구분)",
                              placeholder="예: O=0, -= , _=")
    o1, o2, o3 = st.columns(3)
    upper = o1.checkbox("소문자 → 대문자", value=True)
    no_space = o2.checkbox("공백 모두 제거", value=True)
    fix_mac = o3.checkbox("MAC주소 AA:BB:CC 형식 변환", value=True, disabled="MAC주소" not in df.columns)

    rules = lib.parse_rules(rule_text)
    if rules:
        st.caption("적용할 치환: " + ", ".join(f"'{o}' → '{n}'" for o, n in rules))

    if st.button("🧹 청소 실행", type="primary"):
        if "시리얼번호" not in df.columns:
            st.error("'시리얼번호' 컬럼이 없습니다.")
        else:
            st.session_state.cleaned = lib.clean_data(df, upper, no_space, rules, fix_mac)

    if st.session_state.get("cleaned") is not None:
        cleaned, sn_n, mac_n = st.session_state.cleaned
        st.success(f"시리얼번호 {sn_n}건, MAC주소 {mac_n}건을 수정했습니다.")

        compare = pd.DataFrame({"지점": df["지점"], "자산번호": df.get("자산번호"),
                                "변경 전": df["시리얼번호"], "변경 후": cleaned["시리얼번호"]})
        compare = compare[compare["변경 전"].fillna("").astype(str) != compare["변경 후"]]
        with st.expander(f"시리얼번호 변경 내역 ({len(compare)}건)", expanded=True):
            st.dataframe(compare, width="stretch", hide_index=True)

        if "MAC_오류" in cleaned.columns and (cleaned["MAC_오류"] != "").any():
            bad = cleaned[cleaned["MAC_오류"] != ""]
            st.warning(f"MAC주소 {len(bad)}건은 형식 변환이 불가능해 원래 값을 유지했습니다.")
            st.dataframe(bad[["지점", "자산번호", "MAC주소", "MAC_오류"]], width="stretch", hide_index=True)

        st.download_button("📥 청소된 엑셀 다운로드",
                           data=lib.to_excel_bytes({"청소결과": cleaned, "변경내역": compare}),
                           file_name=f"청소결과_{TODAY:%Y%m%d}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ---------------------------------------------------------------- 하단: 시각화 차트
# 차트 색상 (라이트/다크 모드별로 따로 지정, 색약 구분 검증된 팔레트)
IS_DARK = getattr(st.context.theme, "type", None) == "dark"
SURFACE = "#0e1117" if IS_DARK else "#ffffff"
SERIES = ["#3987e5", "#d95926", "#199e70"] if IS_DARK else ["#2a78d6", "#eb6834", "#1baf7a"]
TEXT_2 = "#c3c2b7" if IS_DARK else "#52514e"
GRID = "rgba(255,255,255,0.08)" if IS_DARK else "rgba(0,0,0,0.08)"
CHART_LAYOUT = dict(height=420, margin=dict(l=10, r=30, t=40, b=20),
                    paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                    font=dict(family="Malgun Gothic, sans-serif", color=TEXT_2, size=13),
                    hoverlabel=dict(font_family="Malgun Gothic, sans-serif"))

st.divider()
st.subheader("📈 시각화")
left, right = st.columns(2)

with left:
    st.markdown("**부서별 기기 보유량**")
    if "부서" in df.columns:
        dept = df["부서"].fillna("").astype(str).str.strip()
        in_stock = int((dept == "재고").sum())
        counts = dept[(dept != "") & (dept != "재고")].value_counts().sort_values()
        fig = go.Figure(go.Bar(
            x=counts.values, y=counts.index, orientation="h",
            marker=dict(color=SERIES[0], cornerradius=4, line=dict(width=0)),
            text=counts.values, textposition="outside", cliponaxis=False,
            textfont=dict(color=TEXT_2),
            hovertemplate="%{y}: <b>%{x}대</b><extra></extra>",
        ))
        fig.update_layout(**CHART_LAYOUT, bargap=0.35, showlegend=False)
        fig.update_xaxes(title="보유 대수", gridcolor=GRID, zeroline=False, automargin=True)
        fig.update_yaxes(title=None, automargin=True, ticksuffix="  ")
        st.plotly_chart(fig, width="stretch", theme=None)
        if in_stock:
            st.caption(f"사용자 미배정 재고 {in_stock}대는 제외했습니다.")
    else:
        st.info("'부서' 컬럼이 없어 표시할 수 없습니다.")

with right:
    st.markdown("**기기 종류별 비율**")
    if "기기종류" in df.columns:
        kinds = df["기기종류"].dropna().astype(str).str.strip().value_counts()
        order = [k for k in ["노트북", "모니터", "태블릿"] if k in kinds.index]
        order += [k for k in kinds.index if k not in order]
        kinds = kinds.reindex(order)
        # 종류 ↔ 색상 고정 (필터가 바뀌어도 같은 종류는 같은 색)
        colors = [SERIES[i] if i < len(SERIES) else "#8a8984" for i in range(len(kinds))]
        fig = go.Figure(go.Pie(
            labels=kinds.index, values=kinds.values, hole=0.45, sort=False, direction="clockwise",
            marker=dict(colors=colors, line=dict(color=SURFACE, width=2)),
            textinfo="label+percent", textposition="outside",
            hovertemplate="%{label}: <b>%{value}대</b> (%{percent})<extra></extra>",
        ))
        fig.update_layout(**CHART_LAYOUT, showlegend=True,
                          legend=dict(orientation="h", y=-0.08, x=0.5, xanchor="center"),
                          annotations=[dict(text=f"총 {int(kinds.sum()):,}대", showarrow=False,
                                            font=dict(size=16, color=TEXT_2))])
        st.plotly_chart(fig, width="stretch", theme=None)
    else:
        st.info("'기기종류' 컬럼이 없어 표시할 수 없습니다.")
