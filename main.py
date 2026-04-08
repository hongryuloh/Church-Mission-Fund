import streamlit as st
import pandas as pd
import io
import json
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
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
@st.cache_data(ttl=300) # 5분마다 갱신
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
        
        # 원본 엑셀 형태 보존 (집계표 덮어쓰기용)
        raw_excel = file_stream.getvalue() 
        
        df_income = pd.read_excel(io.BytesIO(raw_excel), sheet_name='헌금수입')
        df_target = pd.read_excel(io.BytesIO(raw_excel), sheet_name='작정액')
        return df_income, df_target, raw_excel
    except Exception as e:
        st.error(f"데이터 로드 중 오류 발생: {e}")
        return None, None, None

# --- 3. 데이터 계산 로직 (기존 VBA와 동일) ---
def calculate_details(user_name, df_income, df_target, start_year=2026):
    user_info = df_target[df_target.iloc[:, 0].astype(str).str.strip() == user_name]
    if user_info.empty: return None
    
    # 작정액 (빈칸이면 0으로 처리)
    monthly_commit = user_info.iloc[0, 2]
    monthly_commit = float(monthly_commit) if pd.notna(monthly_commit) else 0.0
    
    user_income = df_income[df_income.iloc[:, 2].astype(str).str.strip() == user_name].copy()
    
    # YYYYMM 텍스트 클리닝 (202601.0 -> 202601 변환)
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
                pass 
    return {"name": user_name, "commit": monthly_commit, "alloc": alloc[1:], "labs": lab[1:], "total": sum(monthly_paid.values())}

# --- 4. 엑셀 파일(개인별 집계 시트) 생성 함수 ---
def generate_summary_excel(df_income, df_target, raw_excel, start_year=2026):
    wb = openpyxl.load_workbook(io.BytesIO(raw_excel))
    
    # '개인별 집계' 시트 초기화 또는 생성
    sheet_name = '개인별 집계'
    if sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        ws.delete_rows(1, ws.max_row) 
    else:
        ws = wb.create_sheet(sheet_name)
        
    headers = ["성명", "작정액", "1월", "2월", "3월", "4월", "5월", "6월", "7월", "8월", "9월", "10월", "11월", "12월", "총납부액", "잔액"]
    ws.append(headers)
    
    # 스타일 세팅
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    header_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border
        cell.fill = header_fill

    names_raw = df_target.iloc[1:, 0].dropna().astype(str).str.strip().unique().tolist()
    names = [n for n in names_raw if n and n.lower() != 'nan']
    
    for name in names:
        res = calculate_details(name, df_income, df_target, start_year)
        if res:
            row = [res["name"], res["commit"]]
            for i in range(12): row.append(res["alloc"][i])
            row.append(res["total"])
            balance = max(0, res["commit"] - res["total"]) if res["commit"] > 0 else 0
            row.append(balance)
            ws.append(row)
            
            for col_idx, cell in enumerate(ws[ws.max_row], 1):
                cell.border = thin_border
                if col_idx == 1: cell.alignment = Alignment(horizontal="center")
                else: cell.number_format = '#,##0' 
                    
    # 열 너비 조정
    ws.column_dimensions['A'].width = 12
    for col in ['B', 'O', 'P']: ws.column_dimensions[col].width = 14
    for col in ['C','D','E','F','G','H','I','J','K','L','M','N']: ws.column_dimensions[col].width = 11

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

# --- 5. 앱 화면 구성 ---
st.set_page_config(page_title="2026 선교헌금 관리", layout="wide")
st.title("⛪ 2026 선교헌금 실시간 관리")

with st.spinner('구글 드라이브에서 데이터를 불러오고 있습니다...'):
    df_income, df_target, raw_excel = load_data(FILE_ID)

if df_income is not None and df_target is not None:
    menu = st.sidebar.radio("메뉴", ["🔍 개인별 조회", "🖨️ 전체 집계표"])

    if menu == "🔍 개인별 조회":
        names_raw = df_target.iloc[1:, 0].dropna().astype(str).str.strip().unique().tolist()
        names = [n for n in names_raw if n and n.lower() != 'nan'] 
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
                            
    elif menu == "🖨️ 전체 집계표":
        st.subheader("🖨️ 개인별 집계 시트 업데이트 및 다운로드")
        st.write("기존 엑셀 원본 파일에 **'개인별 집계'** 시트를 최신 내용으로 덮어씌워 생성합니다. 다운로드 후 바로 인쇄하실 수 있습니다.")
        
        with st.spinner("엑셀 파일을 생성 중입니다..."):
            excel_data = generate_summary_excel(df_income, df_target, raw_excel)
            
        st.download_button(
            label="📥 인쇄용 엑셀 파일 다운로드",
            data=excel_data,
            file_name="2026_선교헌금_최종집계.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
