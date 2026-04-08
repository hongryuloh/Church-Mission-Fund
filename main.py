import streamlit as st
import pandas as pd
import io
import json
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload

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
            df_expense = pd.read_excel(io.BytesIO(raw_excel), sheet_name='지출', header=None)
            if not df_expense.empty:
                first_val = str(df_expense.iloc[0, 0]).strip()
                if any(word in first_val for word in ['일자', '날짜', '지출', '일']):
                    df_expense = df_expense[1:].reset_index(drop=True)
                cols = ['일자', '항목', '금액', '비고']
                current_cols = len(df_expense.columns)
                df_expense.columns = cols[:current_cols] + list(df_expense.columns)[current_cols:]
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

def format_date_str(x):
    if pd.isna(x): return ""
    if isinstance(x, pd.Timestamp) or hasattr(x, 'strftime'):
        return x.strftime('%Y-%m-%d')
    x_str = str(x).strip().split(' ')[0].replace('.0', '')
    if len(x_str) == 8 and x_str.isdigit():
        return f"{x_str[:4]}-{x_str[4:6]}-{x_str[6:8]}"
    if len(x_str) == 6 and x_str.isdigit():
        return f"{x_str[:4]}-{x_str[4:6]}-01"
    return x_str

# --- 3. 데이터 계산 로직 ---
def calculate_details(user_name, df_income, df_target, start_year=2026):
    user_info = df_target[df_target.iloc[:, 0].astype(str).str.strip() == user_name]
    if user_info.empty: return None
    
    monthly_commit = user_info.iloc[0, 2]
    monthly_commit = float(monthly_commit) if pd.notna(monthly_commit) else 0.0
    
    user_income = df_income[df_income.iloc[:, 2].astype(str).str.strip() == user_name].copy()
    user_income['YYYYMM_STR'] = user_income.iloc[:, 1].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
    
    # 금액 합계 계산 (iloc 대신 열 인덱스 3 사용)
    monthly_paid = user_income.groupby('YYYYMM_STR').apply(lambda x: x.iloc[:, 3].sum()).to_dict()
    
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
        ws = wb[sheet_name]; ws.delete_rows(1, ws.max_row) 
    else: ws = wb.create_sheet(sheet_name)
    headers = ["성명", "작정액", "1월", "2월", "3월", "4월", "5월", "6월", "7월", "8월", "9월", "10월", "11월", "12월", "총납부액", "잔액"]
    ws.append(headers)
    thin_border, header_fill = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin')), PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    for cell in ws[1]:
        cell.font, cell.alignment, cell.border, cell.fill = Font(bold=True), Alignment(horizontal="center", vertical="center"), thin_border, header_fill
    names_raw = df_target.iloc[1:, 0].dropna().astype(str).str.strip().unique().tolist()
    names = [n for n in names_raw if n and n.lower() != 'nan']
    for name in names:
        res = calculate_details(name, df_income, df_target, start_year)
        if res:
            row = [res["name"], res["commit"]] + res["alloc"] + [res["total"], max(0, res["commit"] - res["total"]) if res["commit"] > 0 else 0]
            ws.append(row)
            for col_idx, cell in enumerate(ws[ws.max_row], 1):
                cell.border = thin_border
                if col_idx == 1: cell.alignment = Alignment(horizontal="center")
                else: cell.number_format = '#,##0' 
    ws.column_dimensions['A'].width = 12
    for col in ['B', 'O', 'P']: ws.column_dimensions[col].width = 14
    output = io.BytesIO(); wb.save(output); return output.getvalue()

# --- 5. 앱 화면 구성 ---
st.set_page_config(page_title="2026 선교헌금 관리", layout="wide")
st.title("⛪ 2026 선교헌금 관리 시스템")
for key in ['add_mode_tgt', 'add_mode_inc', 'add_mode_exp']:
    if key not in st.session_state: st.session_state[key] = False

with st.spinner('데이터를 동기화하는 중입니다...'):
    df_income, df_target, df_expense, raw_excel = load_data(FILE_ID)

if df_income is not None:
    menu = st.sidebar.radio("메뉴", ["🔍 개인별 조회", "✍️ 데이터 관리(입력/조회)", "📊 결산 및 통계", "🖨️ 인쇄용 집계표"])

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

    elif menu == "✍️ 데이터 관리(입력/조회)":
        st.subheader("✍️ 내역 관리 및 입력")
        tab1, tab2, tab3 = st.tabs(["헌금 수입 관리", "지출 내역 관리", "작정액 관리"])
        
        with tab1: # 헌금 수입
            if not st.session_state.add_mode_inc:
                st.write("🔹 최근 헌금 수입 내역")
                # [에러 수정] 열 개수에 맞춰 유동적으로 인덱스 선택
                cols_to_show = [0, 2, 3] # 날짜, 이름, 금액은 필수
                if len(df_income.columns) >= 5: cols_to_show.append(4) # 비고가 있으면 포함
                
                display_inc = df_income.iloc[1:, cols_to_show].copy()
                display_inc.columns = ['헌금일자', '성명', '금액', '비고'][:len(cols_to_show)]
                display_inc['헌금일자'] = display_inc['헌금일자'].apply(format_date_str)
                display_inc['금액'] = pd.to_numeric(display_inc['금액'], errors='coerce').fillna(0).apply(lambda x: f"{int(x):,} 원")
                st.dataframe(display_inc.dropna(subset=['성명']), use_container_width=True, hide_index=True)
                if st.button("➕ 신규 헌금 등록"): st.session_state.add_mode_inc = True; st.rerun()
            else:
                with st.form("inc_form"):
                    d, amt = st.date_input("입금일자"), st.number_input("금액", min_value=0, step=10000)
                    target_data = df_target.iloc[1:, 0:2].copy()
                    target_data.columns = ['N', 'P']
                    options = [f"{r['N']} ({r['P']})" if r['P'] and str(r['P'])!='nan' else str(r['N']) for _, r in target_data.iterrows() if str(r['N'])!='nan']
                    sel = st.selectbox("성명 선택", options); note = st.text_input("비고")
                    if st.form_submit_button("저장") and sel and amt > 0:
                        real_n = sel.split(" (")[0]
                        new = [d.strftime("%Y-%m-%d"), d.strftime("%Y%m"), real_n, amt, note]
                        if save_to_drive(FILE_ID, append_row_to_excel(raw_excel, '헌금수입', new)):
                            st.session_state.add_mode_inc = False; st.rerun()
                    if st.form_submit_button("목록으로 돌아가기"): st.session_state.add_mode_inc = False; st.rerun()

        with tab2: # 지출 내역
            if not st.session_state.add_mode_exp:
                st.write("🔹 지출 내역")
                display_exp = df_expense.copy()
                if '일자' in display_exp.columns: display_exp['일자'] = display_exp['일자'].apply(format_date_str)
                if '금액' in display_exp.columns:
                    display_exp['금액'] = pd.to_numeric(display_exp['금액'], errors='coerce').fillna(0).apply(lambda x: f"{int(x):,} 원")
                st.dataframe(display_exp, use_container_width=True, hide_index=True)
                if st.button("➕ 신규 지출 등록"): st.session_state.add_mode_exp = True; st.rerun()
            else:
                with st.form("exp_form"):
                    d, item, amt, note = st.date_input("지출일자"), st.text_input("항목"), st.number_input("금액", min_value=0, step=10000), st.text_input("비고")
                    if st.form_submit_button("저장") and item and amt > 0:
                        if save_to_drive(FILE_ID, append_row_to_excel(raw_excel, '지출', [d.strftime("%Y-%m-%d"), item, amt, note])):
                            st.session_state.add_mode_exp = False; st.rerun()
                    if st.form_submit_button("목록으로 돌아가기"): st.session_state.add_mode_exp = False; st.rerun()

        with tab3: # 작정액 관리
            if not st.session_state.add_mode_tgt:
                st.write("🔹 작정 명단")
                display_tgt = df_target.iloc[1:, :3].copy()
                display_tgt.columns = ['성명', '직분', '월 작정액']
                display_tgt['월 작정액'] = pd.to_numeric(display_tgt['월 작정액'], errors='coerce').fillna(0).apply(lambda x: f"{int(x):,} 원")
                st.dataframe(display_tgt.dropna(subset=['성명']), use_container_width=True, hide_index=True)
                if st.button("➕ 신규 성도 등록"): st.session_state.add_mode_tgt = True; st.rerun()
            else:
                with st.form("tgt_form"):
                    n, p, amt = st.text_input("성명"), st.text_input("직분"), st.number_input("월 작정액", min_value=0, step=10000)
                    if st.form_submit_button("저장") and n and amt > 0:
                        if save_to_drive(FILE_ID, append_row_to_excel(raw_excel, '작정액', [n, p, amt])):
                            st.session_state.add_mode_tgt = False; st.rerun()
                    if st.form_submit_button("목록으로 돌아가기"): st.session_state.add_mode_tgt = False; st.rerun()

    elif menu == "📊 결산 및 통계":
        st.subheader("📊 재정 결산 및 통계")
        total_inc = pd.to_numeric(df_income.iloc[1:, 3], errors='coerce').sum()
        total_exp = pd.to_numeric(df_expense['금액'], errors='coerce').sum() if not df_expense.empty and '금액' in df_expense.columns else 0
        c1, c2, c3 = st.columns(3)
        c1.metric("총 헌금 수입", f"{int(total_inc):,} 원"); c2.metric("총 지출액", f"{int(total_exp):,} 원"); c3.metric("현재 잔액", f"{int(total_inc - total_exp):,} 원")
        st.divider(); st.write("최근 헌금 5건")
        recent = df_income.tail(5).iloc[:, [0, 2, 3]].copy()
        recent.columns = ['헌금일자', '성명', '금액']
        recent['헌금일자'] = recent['헌금일자'].apply(format_date_str)
        recent['금액'] = pd.to_numeric(recent['금액'], errors='coerce').fillna(0).apply(lambda x: f"{int(x):,} 원")
        st.dataframe(recent, use_container_width=True, hide_index=True)

    elif menu == "🖨️ 인쇄용 집계표":
        st.subheader("🖨️ 개인별 집계 시트 업데이트")
        with st.spinner("엑셀 생성 중..."): data = generate_summary_excel(df_income, df_target, raw_excel)
        st.download_button("📥 인쇄용 엑셀 다운로드", data=data, file_name="2026_선교헌금_최종집계.xlsx")
