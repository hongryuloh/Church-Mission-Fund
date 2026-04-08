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
@st.cache_data(ttl=60) # 데이터 확인을 위해 캐시 시간을 1분으로 단축
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
            cols = ['날짜', '년월', '내역', '금액']
            df_expense.columns = cols[:len(df_expense.columns)] + list(df_expense.columns)[len(cols):]
        except:
            df_expense = pd.DataFrame(columns=['날짜', '년월', '내역', '금액'])
            
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

# [핵심 수정] 실제 데이터가 있는 마지막 행을 찾아 그 다음 행을 반환하는 함수
def find_real_last_row(ws, col_index):
    # col_index: 1(A열), 2(B열), 3(C열) ...
    for row in range(ws.max_row, 0, -1):
        if ws.cell(row=row, column=col_index).value is not None:
            return row + 1
    return 1

def append_row_to_excel(raw_excel, sheet_name, row_data):
    wb = openpyxl.load_workbook(io.BytesIO(raw_excel))
    if sheet_name not in wb.sheetnames:
        ws = wb.create_sheet(sheet_name)
    else:
        ws = wb[sheet_name]
    
    # 성명(C열)이나 내역(C열)이 비어있는 첫 번째 줄 찾기 (3번 열 기준)
    target_row = find_real_last_row(ws, 3)
    
    # 찾은 줄에 데이터 써넣기
    for col, value in enumerate(row_data, 1):
        ws.cell(row=target_row, column=col, value=value)
        
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
    monthly_commit = float(user_info.iloc[0, 2]) if pd.notna(user_info.iloc[0, 2]) else 0.0
    user_income = df_income[df_income.iloc[:, 2].astype(str).str.strip() == user_name].copy()
    user_income['YYYYMM_STR'] = user_income.iloc[:, 1].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
    monthly_paid = user_income.groupby('YYYYMM_STR').apply(lambda x: pd.to_numeric(x.iloc[:, 3], errors='coerce').sum()).to_dict()
    
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
                if 1 <= m <= 12: alloc[m], lab[m] = monthly_paid[pm], f"{m}월납"
            except: pass 
    return {"name": user_name, "commit": monthly_commit, "alloc": alloc[1:], "labs": lab[1:], "total": sum(monthly_paid.values())}

# --- 4. 집계표 생성 ---
def generate_summary_excel(df_income, df_target, raw_excel, start_year=2026):
    wb = openpyxl.load_workbook(io.BytesIO(raw_excel))
    ws = wb['개인별 집계'] if '개인별 집계' in wb.sheetnames else wb.create_sheet('개인별 집계')
    ws.delete_rows(1, ws.max_row)
    headers = ["성명", "작정액", "1월", "2월", "3월", "4월", "5월", "6월", "7월", "8월", "9월", "10월", "11월", "12월", "총납부액", "잔액"]
    ws.append(headers)
    names = [n for n in df_target.iloc[1:, 0].dropna().astype(str).str.strip().unique().tolist() if n and n.lower() != 'nan']
    for n in names:
        res = calculate_details(n, df_income, df_target, start_year)
        if res:
            ws.append([res["name"], res["commit"]] + res["alloc"] + [res["total"], max(0, res["commit"] - res["total"])])
    output = io.BytesIO(); wb.save(output); return output.getvalue()

# --- 5. 화면 구성 ---
st.set_page_config(page_title="2026 선교헌금 관리", layout="wide")
st.title("⛪ 2026 선교헌금 관리 시스템")
for k in ['add_mode_tgt', 'add_mode_inc', 'add_mode_exp']:
    if k not in st.session_state: st.session_state[k] = False

with st.spinner('동기화 중...'):
    df_income, df_target, df_expense, raw_excel = load_data(FILE_ID)

if df_income is not None:
    menu = st.sidebar.radio("메뉴", ["🔍 개인별 조회", "✍️ 데이터 관리(입력/조회)", "📊 결산 및 통계", "🖨️ 인쇄용 집계표"])

    if menu == "🔍 개인별 조회":
        names = [n for n in df_target.iloc[1:, 0].dropna().astype(str).str.strip().unique().tolist() if n and n.lower() != 'nan']
        selected = st.selectbox("성명을 선택하세요", names)
        if selected:
            res = calculate_details(selected, df_income, df_target)
            if res:
                st.subheader(f"📄 {res['name']} 성도님")
                st.write(f"작정액: {int(res['commit']):,}원 | 총납부: {int(res['total']):,}원")
                for r in [0, 6]:
                    cols = st.columns(6)
                    for i in range(6):
                        m, with_col = r + i, cols[i]
                        with with_col:
                            st.write(f"**{m+1}월**"); st.markdown(f"<small>{res['labs'][m]}</small>", unsafe_allow_html=True); st.write(f"{int(res['alloc'][m]):,}원")

    elif menu == "✍️ 데이터 관리(입력/조회)":
        tab1, tab2, tab3 = st.tabs(["헌금 수입", "지출 내역", "작정액 관리"])
        with tab1:
            if not st.session_state.add_mode_inc:
                display_inc = df_income.iloc[1:, [0, 2, 3]].copy(); display_inc.columns = ['날짜', '성명', '금액']
                display_inc['날짜'] = display_inc['날짜'].apply(format_date_str)
                display_inc['금액'] = pd.to_numeric(display_inc['금액'], errors='coerce').fillna(0).apply(lambda x: f"{int(x):,} 원")
                st.dataframe(display_inc.dropna(subset=['성명']), use_container_width=True, hide_index=True)
                if st.button("➕ 신규 헌금 등록"): st.session_state.add_mode_inc = True; st.rerun()
            else:
                with st.form("inc"):
                    d, amt = st.date_input("입금일자"), st.number_input("금액", min_value=0, step=10000)
                    options = [f"{r[0]} ({r[1]})" if pd.notna(r[1]) else str(r[0]) for r in df_target.iloc[1:, 0:2].values if pd.notna(r[0])]
                    sel, note = st.selectbox("성명 선택", options), st.text_input("비고")
                    if st.form_submit_button("저장") and sel and amt > 0:
                        if save_to_drive(FILE_ID, append_row_to_excel(raw_excel, '헌금수입', [d.strftime("%Y-%m-%d"), d.strftime("%Y%m"), sel.split(" (")[0], amt, note])):
                            st.session_state.add_mode_inc = False; st.rerun()
                    if st.form_submit_button("돌아가기"): st.session_state.add_mode_inc = False; st.rerun()

        with tab2:
            if not st.session_state.add_mode_exp:
                display_exp = df_expense.iloc[:, [0, 2, 3]].copy(); display_exp.columns = ['날짜', '내역', '금액']
                display_exp['날짜'] = display_exp['날짜'].apply(format_date_str)
                display_exp['금액'] = pd.to_numeric(display_exp['금액'], errors='coerce').fillna(0).apply(lambda x: f"{int(x):,} 원")
                st.dataframe(display_exp, use_container_width=True, hide_index=True)
                if st.button("➕ 신규 지출 등록"): st.session_state.add_mode_exp = True; st.rerun()
            else:
                with st.form("exp"):
                    d, item, amt, note = st.date_input("지출일자"), st.text_input("항목"), st.number_input("금액", min_value=0, step=10000), st.text_input("비고")
                    if st.form_submit_button("저장") and item and amt > 0:
                        if save_to_drive(FILE_ID, append_row_to_excel(raw_excel, '지출', [d.strftime("%Y-%m-%d"), d.strftime("%Y%m"), item, amt, note])):
                            st.session_state.add_mode_exp = False; st.rerun()
                    if st.form_submit_button("돌아가기"): st.session_state.add_mode_exp = False; st.rerun()

        with tab3:
            if not st.session_state.add_mode_tgt:
                display_tgt = df_target.iloc[1:, :3].copy(); display_tgt.columns = ['성명', '직분', '월 작정액']
                display_tgt['월 작정액'] = pd.to_numeric(display_tgt['월 작정액'], errors='coerce').fillna(0).apply(lambda x: f"{int(x):,} 원")
                st.dataframe(display_tgt.dropna(subset=['성명']), use_container_width=True, hide_index=True)
                if st.button("➕ 신규 성도 등록"): st.session_state.add_mode_tgt = True; st.rerun()
            else:
                with st.form("tgt"):
                    n, p, amt = st.text_input("성명"), st.text_input("직분"), st.number_input("월 작정액", min_value=0, step=10000)
                    if st.form_submit_button("저장") and n and amt > 0:
                        if save_to_drive(FILE_ID, append_row_to_excel(raw_excel, '작정액', [n, p, amt])):
                            st.session_state.add_mode_tgt = False; st.rerun()
                    if st.form_submit_button("돌아가기"): st.session_state.add_mode_tgt = False; st.rerun()

    elif menu == "📊 결산 및 통계":
        t_inc = pd.to_numeric(df_income.iloc[1:, 3], errors='coerce').sum()
        t_exp = pd.to_numeric(df_expense.iloc[:, 3], errors='coerce').sum() if not df_expense.empty else 0
        c1, c2, c3 = st.columns(3)
        c1.metric("총 수입", f"{int(t_inc):,} 원"); c2.metric("총 지출", f"{int(t_exp):,} 원"); c3.metric("현재 잔액", f"{int(t_inc - t_exp):,} 원")

    elif menu == "🖨️ 인쇄용 집계표":
        with st.spinner("생성 중..."): data = generate_summary_excel(df_income, df_target, raw_excel)
        st.download_button("📥 엑셀 다운로드", data=data, file_name="최종집계.xlsx")
