import streamlit as st
import pandas as pd
import msal
import requests
import io

# --- 1. 보안 설정 (Streamlit Secrets에서 가져오기) ---
CLIENT_ID = st.secrets["azure"]["client_id"]
CLIENT_SECRET = st.secrets["azure"]["client_secret"]
TENANT_ID = st.secrets["azure"]["tenant_id"]
FILE_NAME = "2026 선교헌금집계.xlsx" # OneDrive에 있는 정확한 파일명

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPE = ["https://graph.microsoft.com/.default"]

# --- 2. OneDrive 접속용 토큰 받기 함수 ---
def get_access_token():
    app = msal.ConfidentialClientApplication(CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET)
    result = app.acquire_token_for_client(scopes=SCOPE)
    return result.get("access_token")

# --- 3. OneDrive에서 엑셀 파일 다운로드 함수 (경로 지정 버전) ---
def download_excel_from_onedrive():
    token = get_access_token()
    if not token:
        st.error("토큰을 가져오지 못했습니다. Azure 설정을 확인하세요.")
        return None
        
    headers = {'Authorization': f'Bearer {token}'}
    
    # 사용자가 알려주신 폴더 경로와 파일명
    # 경로: wooridongnaechurch / kimhyuncheol - 2026 / 2026 선교헌금집계.xlsx
    folder_path = "wooridongnaechurch/kimhyuncheol - 2026"
    file_name = "2026 선교헌금집계.xlsx"
    
    # Microsoft Graph API 경로 방식 URL (공백은 자동으로 처리됨)
    url = f"https://graph.microsoft.com/v1.0/me/drive/root:/{folder_path}/{file_name}"
    
    response = requests.get(url, headers=headers)
    res_json = response.json()
    
    if '@microsoft.graph.downloadUrl' in res_json:
        download_url = res_json['@microsoft.graph.downloadUrl']
        file_content = requests.get(download_url).content
        return io.BytesIO(file_content)
    else:
        # 에러 메시지 상세 출력 (디버깅용)
        error_msg = res_json.get('error', {}).get('message', '파일을 찾을 수 없습니다.')
        st.error(f"OneDrive 오류: {error_msg}")
        st.info(f"찾으려는 경로: {folder_path}/{file_name}")
        return None

# --- 4. 데이터 계산 로직 (VBA 로직 재현) ---
def calculate_details(user_name, df_income, df_target, start_year=2026):
    user_info = df_target[df_target.iloc[:, 0].astype(str).str.strip() == user_name]
    if user_info.empty: return None
    monthly_commit = float(user_info.iloc[0, 2])
    user_income = df_income[df_income.iloc[:, 2].astype(str).str.strip() == user_name].copy()
    user_income['YYYYMM_STR'] = user_income.iloc[:, 1].astype(str).str.strip()
    monthly_paid = user_income.groupby('YYYYMM_STR').iloc[:, 3].sum().to_dict()
    
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
                    txt = f"{pm[-2:]}월납"
                    lab[ptr] = txt if lab[ptr] == "" else f"{lab[ptr]}<br>{txt}"
                    alloc[ptr] += amt
                    val -= amt
                    if alloc[ptr] >= monthly_commit - 0.0001: ptr += 1
    else:
        for pm in sorted_months:
            m = int(pm[-2:])
            if 1 <= m <= 12: alloc[m], lab[m] = monthly_paid[pm], f"{pm[-2:]}월납"
    return {"name": user_name, "commit": monthly_commit, "alloc": alloc[1:], "labs": lab[1:], "total": sum(monthly_paid.values())}

# --- 5. 앱 화면 구성 ---
st.set_page_config(page_title="2026 선교헌금 관리", layout="wide")
st.title("⛪ 2026 선교헌금 실시간 관리")

if 'df_income' not in st.session_state:
    with st.spinner('OneDrive에서 데이터를 불러오고 있습니다...'):
        file_bytes = download_excel_from_onedrive()
        if file_bytes:
            st.session_state.df_income = pd.read_excel(file_bytes, sheet_name='헌금수입')
            st.session_state.df_target = pd.read_excel(file_bytes, sheet_name='작정액')
            st.success("데이터 로드 완료!")
        else:
            st.error(f"OneDrive에서 '{FILE_NAME}' 파일을 찾을 수 없습니다.")

if 'df_income' in st.session_state:
    menu = st.sidebar.radio("메뉴", ["🔍 개인별 조회", "🖨️ 전체 집계표"])
    df_income, df_target = st.session_state.df_income, st.session_state.df_target

    if menu == "🔍 개인별 조회":
        names = df_target.iloc[2:, 0].dropna().unique().tolist()
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
