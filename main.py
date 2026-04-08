import streamlit as st
import pandas as pd
import io
import json
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload
from datetime import datetime

# --- 1. 보안 설정 ---
FILE_ID = st.secrets["google"]["file_id"]

def get_gdrive_service():
    creds_json = json.loads(st.secrets["google"]["service_account"])
    credentials = service_account.Credentials.from_service_account_info(
        creds_json, scopes=['https://www.googleapis.com/auth/drive']
    )
    return build('drive', 'v3', credentials=credentials)

# --- 2. 구글 드라이브 데이터 로드 & 저장 함수 ---
@st.cache_data(ttl=300) 
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
        
        raw_excel = file_stream.getvalue() 
        df_income = pd.read_excel(io.BytesIO(raw_excel), sheet_name='헌금수입')
        df_target = pd.read_excel(io.BytesIO(raw_excel), sheet_name='작정액')
        
        try:
            df_expense = pd.read_excel(io.BytesIO(raw_excel), sheet_name='지출')
        except:
            df_expense = pd.DataFrame(columns=['일자', '항목', '금액', '비고'])
            
        return df_income, df_target, df_expense, raw_excel
    except Exception as e:
        st.error(f"데이터 로드 중 오류 발생: {e}")
        return None, None, None, None

def save_to_drive(file_id, excel_bytes):
    try:
        service = get_gdrive_service()
        media = MediaIoBaseUpload(io.BytesIO(excel_bytes), mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        service.files().update(fileId=file_id, media_body=media).execute()
        st.cache_data.clear() 
        return True
    except Exception as e:
        st.error(f"저장 실패: {e}")
        return False

def append_row_to_excel(raw_excel, sheet_name, row_data):
    wb = openpyxl.load_workbook(io.BytesIO(raw_excel))
    if sheet_name not in wb.sheetnames:
        ws = wb.create_sheet(sheet_name)
    else:
        ws = wb[sheet_name]
    
    ws.append(row_data)
    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

# --- 3. 데이터 계산 로직 ---
def calculate_details(user_name, df_income, df_target, start_year=2026):
    user_info = df_target[df_target.iloc[:, 0].astype(str).str.strip() == user_name]
    if user_info.empty: return None
    
    monthly_commit = user_info.iloc[0, 2]
    monthly_commit = float(monthly_commit) if pd.notna(monthly_commit) else 0.0
    
    user_income = df_income[df_income.iloc[:, 2].astype(str).str.strip() == user_name].copy()
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
            except: pass 
    return {"name": user_name, "commit": monthly_commit, "alloc": alloc[1:], "labs": lab[1:], "total": sum(monthly_paid.values())}

# --- 4. 집계표 엑셀 생성 함수 ---
def generate_summary_excel(df_income, df_target, raw_excel, start_year=2026):
    wb = openpyxl.load_workbook(io.BytesIO(raw_excel))
    sheet_name = '개인별 집계'
    if sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        ws.delete_rows(1, ws.max_row) 
    else:
        ws = wb.create_sheet(sheet_name)
        
    headers = ["성명", "작정액", "1월", "2월", "3월", "4월", "5월", "6월", "7월", "8월", "9월", "10월", "11월", "12월", "총납부액", "잔액"]
    ws.append(headers)
    
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
                    
    ws.column_dimensions['A'].width = 12
    for col in ['B', 'O', 'P']: ws.column_dimensions[col].width = 14
    for col in ['C','D','E','F','G','H','I','J','K','L','M','N']: ws.column_dimensions[col].width = 11

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

# --- 5. 앱 화면 구성 ---
st.set_page_config(page_title="2026 선교헌금 관리", layout="wide")
st.title("⛪ 2026 선교헌금 관리 시스템")

# 세션 상태 초기화 (입력 모드 전환용)
if 'add_mode' not in st.session_state:
    st.session_state.add_mode = False

with st.spinner('데이터를 동기화하는 중입니다...'):
    df_income, df_target, df_expense, raw_excel = load_data(FILE_ID)

if df_income is not None:
    menu = st.sidebar.radio("메뉴", ["🔍 개인별 조회", "✍️ 데이터 입력", "📊 결산 및 통계", "🖨️ 인쇄용 집계표"])

    # ----------------------------------------------------------------
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
                            
    # ----------------------------------------------------------------
    elif menu == "✍️ 데이터 입력":
        st.subheader("✍️ 내역 입력하기")
        tab1, tab2, tab3 = st.tabs(["헌금 수입", "지출 내역", "작정액 관리"])
        
        with tab1:
            with st.form("income_form"):
                col1, col2 = st.columns(2)
                inc_date = col1.date_input("입금일자")
                inc_name = col2.text_input("성명")
                inc_amt = col1.number_input("금액", min_value=0, step=10000)
                inc_note = col2.text_input("비고 (선택)")
                submitted1 = st.form_submit_button("헌금 수입 저장")
                
                if submitted1 and inc_name and inc_amt > 0:
                    yyyymm = inc_date.strftime("%Y%m")
                    new_row = ["", yyyymm, inc_name, inc_amt, inc_note] 
                    with st.spinner("저장 중..."):
                        updated_excel = append_row_to_excel(raw_excel, '헌금수입', new_row)
                        if save_to_drive(FILE_ID, updated_excel):
                            st.success(f"{inc_name} 성도님의 헌금이 저장되었습니다.")
                            st.rerun()

        with tab2:
            with st.form("expense_form"):
                col1, col2 = st.columns(2)
                exp_date = col1.date_input("지출일자")
                exp_item = col2.text_input("지출항목")
                exp_amt = col1.number_input("지출 금액", min_value=0, step=10000)
                exp_note = col2.text_input("비고 (선택)")
                submitted2 = st.form_submit_button("지출 내역 저장")
                
                if submitted2 and exp_item and exp_amt > 0:
                    new_row = [exp_date.strftime("%Y-%m-%d"), exp_item, exp_amt, exp_note]
                    with st.spinner("저장 중..."):
                        updated_excel = append_row_to_excel(raw_excel, '지출', new_row)
                        if save_to_drive(FILE_ID, updated_excel):
                            st.success(f"지출 내역이 저장되었습니다.")
                            st.rerun()
                            
        with tab3: # 작정액 관리 통합 화면
            if not st.session_state.add_mode:
                st.write(" 현재 등록된 작정 명단")
                # 작정액 시트 정리해서 표시 (1행 제목 제외)
                display_target = df_target.iloc[1:, :3].copy()
                display_target.columns = ['성명', '직분', '월 작정액']
                st.dataframe(display_target.dropna(subset=['성명']), use_container_width=True, hide_index=True)
                
                if st.button("➕ 신규 성도 추가"):
                    st.session_state.add_mode = True
                    st.rerun()
            else:
                st.write(" 신규 성도 작정액 등록")
                with st.form("target_form_new"):
                    new_name = st.text_input("성명")
                    new_pos = st.text_input("직분")
                    new_amt = st.number_input("월 작정액", min_value=0, step=10000)
                    
                    c1, c2 = st.columns(2)
                    save_btn = c1.form_submit_button("저장하기")
                    cancel_btn = c2.form_submit_button("취소")
                    
                    if save_btn and new_name and new_amt > 0:
                        new_row = [new_name, new_pos, new_amt]
                        with st.spinner("저장 중..."):
                            updated_excel = append_row_to_excel(raw_excel, '작정액', new_row)
                            if save_to_drive(FILE_ID, updated_excel):
                                st.session_state.add_mode = False
                                st.success(f"{new_name} 성도님의 정보가 저장되었습니다.")
                                st.rerun()
                    if cancel_btn:
                        st.session_state.add_mode = False
                        st.rerun()

    # ----------------------------------------------------------------
    elif menu == "📊 결산 및 통계":
        st.subheader("📊 재정 결산 및 통계")
        total_income = df_income.iloc[:, 3].sum() if not df_income.empty else 0
        total_expense = df_expense['금액'].sum() if not df_expense.empty else 0
        balance = total_income - total_expense
        
        col1, col2, col3 = st.columns(3)
        col1.metric("총 헌금 수입", f"{int(total_income):,} 원")
        col2.metric("총 지출액", f"{int(total_expense):,} 원")
        col3.metric("현재 잔액", f"{int(balance):,} 원")
        
        st.divider()
        st.write("주단위 헌금 집계 (최근 5건)")
        st.dataframe(df_income.tail(5).iloc[:, [1,2,3,4]], use_container_width=True, hide_index=True)

    # ----------------------------------------------------------------
    elif menu == "🖨️ 인쇄용 집계표":
        st.subheader("🖨️ 개인별 집계 시트 업데이트")
        with st.spinner("엑셀 생성 중..."):
            excel_data = generate_summary_excel(df_income, df_target, raw_excel)
        st.download_button("📥 인쇄용 엑셀 다운로드", data=excel_data, file_name="2026_선교헌금_최종집계.xlsx")
