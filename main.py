import streamlit as st
import psycopg2
import traceback

st.title("🔍 데이터베이스 연결 정밀 진단기")

try:
    # 1. Secrets에서 주소 가져오기
    db_url = st.secrets["connections"]["postgresql"]["url"]
    
    # 보안을 위해 화면에는 비밀번호를 가리고 주소 형태만 출력합니다.
    safe_url = db_url.replace(db_url.split(":")[2].split("@")[0], "********")
    st.info(f"입력된 주소 형태: {safe_url}")
    
    # 2. 강제 직접 연결 시도
    st.write("연결을 시도하는 중...")
    conn = psycopg2.connect(db_url)
    
    # 3. 성공 시
    st.success("🎉 DB 연결에 성공했습니다!! (주소와 비밀번호가 모두 맞습니다)")
    conn.close()

except Exception as e:
    # 4. 실패 시 (숨겨졌던 진짜 에러 메시지 출력)
    st.error("❌ 연결 실패! 아래의 '진짜 에러 원인'을 복사해서 제미나이에게 알려주세요.")
    st.code(traceback.format_exc())
