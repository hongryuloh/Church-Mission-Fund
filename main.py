import streamlit as st
import pandas as pd
import io
import json
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

# --- 1. 보안 설정 (Streamlit Secrets에서 가져오기) ---
FILE_ID = st.secrets["google"]["file_id"]

def get_gdrive_service():
    creds_json = json.loads(st.secrets["google"]["service_account"])
    credentials = service_account.Credentials.from_service_account_info(
        creds_json, scopes=['https://www.googleapis.com/auth/drive.readonly']
    )
    return build('drive', 'v3', credentials=credentials)

# --- 2. 구글 드라이브 엑셀 다운로드 함수 ---
@st.cache_data(ttl=300) # 5분마다 갱신하여 데이터 최신화
def load_data(file_id):
    try:
        service = get_gdrive_service()
        request = service.files().get_media(fileId=file_id)
        file_stream = io.BytesIO()
        downloader = MediaIoBaseDownload(file_stream, request)
        done = False
        while not done:
            status, done = downloader.next_chunk()
        file_stream.seek(0)
        
        # 다운로드한 엑셀 파일 읽기
        df_income = pd.read_excel(file_stream, sheet_name='헌금수입')
        df_target = pd.read_excel(file_stream, sheet_name='작정액')
        return df_income, df_target
    except Exception as e:
        st.error(f"데이터 로드 중 오류 발생: {e}")
        return None, None

# --- 3. 데이터 계산 로직 (기존 VBA와 동일) ---
def calculate_details(user_name, df_income, df_target, start_year=2026):
    user_info = df_target[df_target.iloc[:, 0].astype(str).str.strip() == user_name]
    if user_info.empty: return None
    
    # 작정액 (빈칸이면 0으로 처리)
    monthly_commit = user_info.iloc[0, 2]
    monthly_commit = float(monthly_commit) if pd.notna(monthly_commit) else 0.0
    
    user_income = df_income[df_income.iloc[:, 2].astype(str).str.strip() == user_name].copy()
    
    # 수정 1: YYYYMM 텍스트 클리닝 (202601.0 -> 202601 변환하여 에러 방지)
    user_income['YYYYMM_STR'] = user_income.iloc[:, 1].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
    amt_col_name = user_income.columns[3] 
    monthly_paid = user_income.groupby('YYYYMM_STR')[amt_col_name].sum().to_dict()
    
    alloc, lab = [0.0]*13, [""]*13
    sorted_months = sorted([m for m in monthly_paid.keys() if m.startswith(str(start_year))])

    if monthly_commit > 0:
        ptr = 1
        for pm in sorted_months:
            val = float(monthly_paid[pm])
            while val > 0 and ptr <= 12:
                rem = monthly_commit - alloc[ptr]
                if rem <= 0: ptr += 1
                else:
                    amt = min(val, rem)
                    txt = f"{int(pm[-2:])}월납"
                    lab[ptr] = txt if lab[ptr] == "" else f"{lab[ptr]}<br>{txt}"
                    alloc[ptr] += amt
                    val -= amt
                    if alloc[ptr] >= monthly_commit - 0.0001: ptr += 1
    else:
        for pm in sorted_months:
            try:
                m = int(pm[-2:])
                if 1 <= m <= 12: 
                    alloc[m] = monthly_paid[pm]
                    lab[m] = f"{m}월납"
            except:
                pass # 변환 실패 시 안전하게 무시
    return {"name": user_name, "commit": monthly_commit, "alloc": alloc[1:], "labs": lab[1:], "total": sum(monthly_paid.values())}

# --- 4. 앱 화면 구성 ---
st.set_page_config(page_title="2026 선교헌금 관리", layout="wide")
st.title("⛪ 2026 선교헌금 관리")

with st.spinner('구글 드라이브에서 엑셀 파일을 불러오고 있습니다...'):
    df_income, df_target = load_data(FILE_ID)

if df_income is not None and df_target is not None:
    menu = st.sidebar.radio("메뉴", ["🔍 개인별 조회", "🖨️ 전체 집계표"])

    if menu == "🔍 개인별 조회":
        # 수정 2: iloc[1:, 0]으로 변경하여 엑셀의 3행부터 이름을 가져옴
        names_raw = df_target.iloc[1:, 0].dropna().astype(str).str.strip().unique().tolist()
        names = [n for n in names_raw if n and n.lower() != 'nan'] # 빈칸/에러값 제거
        
        selected = st.selectbox("성명을 선택하세요", names)
        if selected:
            res = calculate_details(selected, df_income, df_target)
            if res:
                st.subheader(f"📄 {res['name']} 성도님")
                st.write(f"작정액: {int(res['commit']):,}원 | 총납부: {int(res['total']):,}원")
                for r in [0, 6]:
                    cols = st.columns(6)
                    for i in range(6):
                        m = r + i
                        with cols[i]:
                            st.write(f"**{m+1}월**")
                            st.markdown(f"<small>{res['labs'][m]}</small>", unsafe_allow_html=True)
                            st.write(f"{int(res['alloc'][m]):,}원")
