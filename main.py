import streamlit as st
import pandas as pd

# 페이지 설정
st.set_page_config(page_title="2026 선교헌금 관리", layout="wide")

# CSS: 인쇄용 점선 및 디자인
st.markdown("""
    <style>
    .dotted-line { border-top: 2px dotted #000; margin: 30px 0; }
    .report-box { border: 1px solid #ccc; padding: 20px; margin-bottom: 10px; background-color: white; }
    </style>
    """, unsafe_allow_html=True)

st.title("⛪ 2026 선교헌금 관리 시스템")
st.info("OneDrive 연결 설정을 완료하면 실시간 데이터가 여기에 나타납니다.")

# 임시 파일 업로더 (연결 전 테스트용)
uploaded_file = st.file_uploader("xlsx 파일을 업로드하세요", type="xlsx")

if uploaded_file:
    st.success("파일이 로드되었습니다. 곧 OneDrive 자동 연동 기능을 활성화하겠습니다!")
