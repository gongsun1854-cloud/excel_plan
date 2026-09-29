"""IT 자산 통합 관리 시스템 (실행: streamlit run app.py)"""
import pandas as pd
import streamlit as st

st.set_page_config(page_title="IT 자산 통합 관리 시스템", page_icon="🖥️", layout="wide")
st.title("🖥️ IT 자산 통합 관리 시스템")

uploaded = st.file_uploader("자산 엑셀 파일을 업로드하세요", type=["xlsx", "xls"])

if uploaded is None:
    st.info("엑셀 파일(.xlsx / .xls)을 올리면 자산 현황이 표시됩니다.")
    st.stop()

try:
    df = pd.read_excel(uploaded)
except Exception as e:
    st.error(f"파일을 읽을 수 없습니다: {e}")
    st.stop()

if "상태" not in df.columns:
    st.error(f"'상태' 컬럼이 없습니다. 파일의 컬럼: {', '.join(map(str, df.columns))}")
    st.stop()

status = df["상태"].astype(str).str.strip()

col1, col2, col3 = st.columns(3)
col1.metric("전체 기기 수", f"{len(df):,}대")
col2.metric("수리 필요 기기 수", f"{(status == '수리필요').sum():,}대")
col3.metric("폐기 예정 기기 수", f"{(status == '폐기예정').sum():,}대")

st.subheader("📋 업로드 데이터 미리보기")
st.caption(f"{uploaded.name} · {len(df):,}행 × {len(df.columns)}열")
st.dataframe(df, width="stretch", hide_index=True)
