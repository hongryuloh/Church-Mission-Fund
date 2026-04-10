import streamlit as st
import pandas as pd
import io
import json
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.worksheet.pagebreak import Break
from openpyxl.utils import get_column_letter
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload
from datetime import datetime

# --- 1. 앱 기본 설정 및 모바일 최적화 CSS ---
# initial_sidebar_state="expanded"를 통해 모바일에서도 메뉴가 먼저 보이도록 유도합니다.
st.set_page_config(page_title="2026 선교헌금 관리", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
    <style>
    /* 1. 상단 여백 제거 및 헤더 숨김 */
    .block-container { padding-top: 0.5rem !important; padding-bottom: 5rem !important; }
    header { visibility: hidden; }
    #MainMenu { visibility: hidden; }
    
    /* 2. 모바일 메뉴 버튼(화살표) 강조 스타일 */
    /* 사이드바가 접혀있을 때 왼쪽 상단에 나타나는 버튼을 강조하여 메뉴 위치를 알립니다. */
    button[kind="headerNoPadding"] {
        background-color: #ff4b4b !important;
        color: white !important;
        border-radius: 50% !important;
        width: 38px !important;
        height: 38px !important;
        top: 8px !important;
        left: 8px !important;
        box-shadow: 0 2px 5px rgba(0,0,0,0.2) !important;
        z-index: 999999;
    }

    /* 3. 모바일에서 버튼 4개를 무조건 한 줄로 가로 배치 */
    div[data-testid="stHorizontalBlock"] {
        display: flex !important;
        flex-direction: row !important;
        flex-wrap: nowrap !important;
        align-items: flex-end !important;
        gap: 5px !important;
    }
    div[data-testid="stHorizontalBlock"] > div {
        flex: 1 1 auto !important;
        min-width: 0 !important;
    }
    
    /* 4. 하단 고정 영역 스타일 (Footer) */
    .fixed-footer {
        position: fixed;
        bottom: 0;
        left: 0;
        width: 100%;
        background-color: white;
        padding: 10px 15px 30px 15px;
        border-top: 1px solid #ddd;
        z-index: 999;
    }
    @media (prefers-color-scheme: dark) {
        .fixed-footer { background-color: #1e1e1e !important; border-top: 1px solid #333 !important; }
    }
    
    /* 5. 입력창 및 버튼 높이 최적화 */
    .stButton button { width: 100% !important; padding: 0px !important; font-size: 13px !important; height: 40px !important; }
    .stNumberInput input { height: 40px !important; }
    
    /* 6. 사이드바 라디오 버튼 텍스트 크기 조정 (모바일 가독성) */
    [data-testid="stSidebar"] .stRadio div[role="radiogroup"] label {
        font-size: 15px !important;
        padding: 5px 0px !important;
    }
    </style>
    """, unsafe_allow_html=True)

# (1) 세션 상태 초기화
if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False
if "current_user" not in st.session_state:
    st.session_state["current_user"] = ""

# (2) 로그인 화면 구성
if not st.session_state["authenticated"]:
    col1, col2, col3 = st.columns([1, 2, 1]) 
    with col2:
        st.title("⛪ 선교헌금 관리 시스템")
        st.info("🔒 접근 권한이 필요합니다. ID와 비밀번호를 입력해 주세요.")
        with st.form("login_form"):
            input_id = st.text_input("아이디 (ID)")
            input_pwd = st.text_input("비밀번호 (Password)", type="password")
            if st.form_submit_button("로그인", use_container_width=True):
                credentials = st.secrets.get("credentials", {})
                if input_id in credentials and credentials[input_id] == input_pwd:
                    st.session_state["authenticated"] = True
                    st.session_state["current_user"] = input_id
                    st.rerun() 
                else:
                    st.error("❌ 아이디 또는 비밀번호가 틀렸습니다.")
    st.stop() 

# [모바일 최적화] 제목과 로그아웃 버튼을 사이드바 상단으로 이동
with st.sidebar:
    st.title("⛪ 2026 선교헌금")
    st.write(f"👤 **{st.session_state['current_user']}**님 접속 중")
    if st.button("로그아웃", use_container_width=True):
        st.session_state["authenticated"] = False
        st.session_state["current_user"] = ""
        st.rerun()
    st.markdown("---") 

# --- 2. 보안 설정 및 헬퍼 함수 ---
FILE_ID = st.secrets["google"]["file_id"]

def get_gdrive_service():
    creds_json = json.loads(st.secrets["google"]["service_account"])
    credentials = service_account.Credentials.from_service_account_info(
        creds_json, scopes=['https://www.googleapis.com/auth/drive']
    )
    return build('drive', 'v3', credentials=credentials)

def clean_str(x):
    return str(x).split('.')[0].strip() if pd.notna(x) else ""

def get_col(df, possible_names, default_idx):
    for name in possible_names:
        if name in df.columns: return name
    if len(df.columns) > default_idx: return df.columns[default_idx]
    return None

def format_date_str(x):
    if pd.isna(x): return ""
    if isinstance(x, pd.Timestamp) or hasattr(x, 'strftime'): return x.strftime('%Y-%m-%d')
    x_str = str(x).strip().split(' ')[0].replace('.0', '')
    if len(x_str) == 8 and x_str.isdigit(): return f"{x_str[:4]}-{x_str[4:6]}-{x_str[6:8]}"
    if len(x_str) == 6 and x_str.isdigit(): return f"{x_str[:4]}-{x_str[4:6]}-01"
    return x_str

def fmt(val):
    if pd.isna(val) or val == 0: return "-"
    return f"{int(val):,}"

@st.cache_data(ttl=60) 
def load_data(file_id):
    try:
        service = get_gdrive_service()
        request = service.files().get_media(fileId=file_id)
        file_stream = io.BytesIO()
        downloader = MediaIoBaseDownload(file_stream, request)
        done = False
        while not done: status, done = downloader.next_chunk()
        file_stream.seek(0)
        raw_excel = file_stream.getvalue() 
        
        def robust_load(sheet_name, default_cols):
            try:
                df_raw = pd.read_excel(io.BytesIO(raw_excel), sheet_name=sheet_name, header=None).astype(object)
                header_idx = 0
                for i in range(min(10, len(df_raw))):
                    vals = [str(x).strip() for x in df_raw.iloc[i].values if pd.notna(x)]
                    if any(k in vals for k in ['날짜', '이름', '성명', '내역', '작정액', '월별 작정액']):
                        header_idx = i; break
                df = pd.read_excel(io.BytesIO(raw_excel), sheet_name=sheet_name, header=header_idx).astype(object)
                return df.dropna(how='all').reset_index(drop=True)
            except: return pd.DataFrame(columns=default_cols).astype(object)

        df_income = robust_load('헌금수입', ['날짜', '년월', '이름', '금액', '비고'])
        df_target = robust_load('작정액', ['이름', '직분', '월별 작정액', '년간작정금액', '헌금액', '년간작정 잔여금액', '인쇄여부'])
        df_expense = robust_load('지출', ['날짜', '년월', '내역', '금액', '비고'])
        
        for df in [df_income, df_expense]:
            ym_col = get_col(df, ['년월'], 1)
            if ym_col:
                df[ym_col] = df[ym_col].apply(clean_str)
                for idx, row in df.iterrows():
                    if str(row[ym_col]).strip() in ['None', 'nan', '']:
                        d_val = format_date_str(row.get('날짜', '')).replace('-', '')
                        if len(d_val) >= 6: df.at[idx, ym_col] = d_val[:6]
                        
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
    except: return False

def overwrite_sheet_preserve(raw_excel, sheet_name, df_new):
    wb = openpyxl.load_workbook(io.BytesIO(raw_excel))
    ws = wb[sheet_name] if sheet_name in wb.sheetnames else wb.create_sheet(sheet_name)
    header_row = 1
    for r in range(1, 11):
        vals = [str(ws.cell(row=r, column=c).value).strip() for c in range(1, ws.max_column + 1) if ws.cell(row=r, column=c).value is not None]
        if any(v in ['날짜', '이름', '성명', '내역', '작정액', '월별 작정액'] for v in vals):
            header_row = r; break
    if ws.max_row > header_row: ws.delete_rows(header_row + 1, ws.max_row - header_row)
    header_map = {str(ws.cell(row=header_row, column=c).value).strip(): c for c in range(1, ws.max_column + 1) if ws.cell(row=header_row, column=c).value is not None}
    target_row = header_row + 1
    for _, row_data in df_new.iterrows():
        for col_name, val in row_data.items():
            col_str = str(col_name).strip()
            if col_str in header_map:
                ws.cell(row=target_row, column=header_map[col_str], value=None if pd.isna(val) else val)
        target_row += 1
    output = io.BytesIO(); wb.save(output); return output.getvalue()

def append_dict_to_excel(raw_excel, sheet_name, row_dict):
    wb = openpyxl.load_workbook(io.BytesIO(raw_excel))
    ws = wb[sheet_name] if sheet_name in wb.sheetnames else wb.create_sheet(sheet_name)
    header_row = 1
    for r in range(1, 11):
        vals = [str(ws.cell(row=r, column=c).value).strip() for c in range(1, ws.max_column + 1) if ws.cell(row=r, column=c).value is not None]
        if any(v in ['날짜', '이름', '성명', '내역', '작정액', '월별 작정액'] for v in vals): header_row = r; break
    header_map = {str(ws.cell(row=header_row, column=c).value).strip(): c for c in range(1, ws.max_column + 1) if ws.cell(row=header_row, column=c).value is not None}
    last_row = header_row
    for row in range(ws.max_row, header_row - 1, -1):
        if any(ws.cell(row=row, column=c).value is not None for c in range(1, ws.max_column + 1)): last_row = row; break
    target_row = last_row + 1
    for key, val in row_dict.items():
        if key not in header_map:
            new_col = ws.max_column + 1
            ws.cell(row=header_row, column=new_col, value=key)
            header_map[key] = new_col
        ws.cell(row=target_row, column=header_map[key], value=val)
    output = io.BytesIO(); wb.save(output); return output.getvalue()

# --- 3. 데이터 계산 ---
def calculate_details(user_name, df_income, df_target, start_year=2026):
    t_n = get_col(df_target, ['이름', '성명'], 0)
    t_a = get_col(df_target, ['월별 작정액', '작정액'], 2)
    i_n = get_col(df_income, ['이름', '성명'], 2)
    i_y = get_col(df_income, ['년월'], 1)
    i_a = get_col(df_income, ['금액'], 3)
    user_info = df_target[df_target[t_n].astype(str).str.strip() == user_name.strip()]
    if user_info.empty: return None
    commit = pd.to_numeric(user_info.iloc[0].get(t_a, 0), errors='coerce')
    commit = 0.0 if pd.isna(commit) else float(commit)
    total_donated = pd.to_numeric(user_info.iloc[0].get('헌금액', 0), errors='coerce')
    total_donated = 0.0 if pd.isna(total_donated) else float(total_donated)
    u_inc = df_income[df_income[i_n].astype(str).str.strip() == user_name.strip()].copy()
    u_inc[i_a] = pd.to_numeric(u_inc[i_a], errors='coerce').fillna(0)
    paid = u_inc.groupby(i_y)[i_a].sum().to_dict()
    alloc, lab = [0.0]*13, [""]*13
    sorted_m = sorted([m for m in paid.keys() if str(m).startswith(str(start_year))])
    if commit > 0:
        ptr = 1
        for pm in sorted_m:
            val = float(paid.get(pm, 0))
            while val > 0.01 and ptr <= 12:
                rem = commit - alloc[ptr]
                if rem <= 0.01: ptr += 1
                else:
                    amt = min(val, rem); txt = f"{int(str(pm)[-2:]):02d}월납"
                    lab[ptr] = txt if lab[ptr] == "" else f"{lab[ptr]}\n{txt}"
                    alloc[ptr] += amt; val -= amt
                    if alloc[ptr] >= commit - 0.01: ptr += 1
    else:
        for pm in sorted_m:
            try:
                m = int(str(pm)[-2:])
                if 1 <= m <= 12: alloc[m] = float(paid.get(pm, 0)); lab[m] = f"{m:02d}월납"
            except: pass 
    return {"name": user_name, "commit": commit, "alloc": alloc[1:], "labs": lab[1:], "total": total_donated}

# --- 4. 인쇄 포맷 ---
def generate_summary_excel(df_income, df_target, target_month, start_year=2026):
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "개인별 헌금내역"
    for c in range(1, 19): ws.column_dimensions[get_column_letter(c)].width = 8.13
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    t_n, t_p = get_col(df_target, ['이름', '성명'], 0), get_col(df_target, ['직분'], 1)
    print_users = df_target[df_target['인쇄여부'] == 'Y'] 
    user_list = []
    for _, row_data in print_users.iterrows():
        name, pos = clean_str(row_data.get(t_n, '')), clean_str(row_data.get(t_p, ''))
        if not name or name == 'nan' or name == '합계': continue
        res = calculate_details(name, df_income, df_target, start_year)
        if res: res['pos'] = pos; user_list.append(res)
    today_str = datetime.now().strftime("%Y.%m.%d")
    current_row = 1
    def draw_user_block(r, c_off, user):
        for r_i in range(r, r+9): ws.row_dimensions[r_i].height = 25 
        ws.merge_cells(start_row=r, start_column=1+c_off, end_row=r, end_column=8+c_off)
        ws.cell(row=r, column=1+c_off, value="2026년 선교헌금 작정 및 헌금내역").font = Font(size=16, bold=True, underline="single")
        ws.cell(row=r, column=1+c_off).alignment = center_align
        ws.merge_cells(start_row=r+1, start_column=1+c_off, end_row=r+1, end_column=3+c_off); ws.cell(row=r+1, column=1+c_off, value=f"({today_str} 기준)").font = Font(bold=True)
        ws.merge_cells(start_row=r+1, start_column=4+c_off, end_row=r+1, end_column=6+c_off); ws.cell(row=r+1, column=4+c_off, value="선교헌금 합계 :").alignment = Alignment(horizontal='right', vertical='center')
        ws.merge_cells(start_row=r+1, start_column=7+c_off, end_row=r+1, end_column=8+c_off); t_c = ws.cell(row=r+1, column=7+c_off, value=user['total']); t_c.font, t_c.number_format, t_c.alignment = Font(bold=True), '#,##0', center_align
        for row_i in range(r+2, r+8):
            for col_i in range(1+c_off, 9+c_off): ws.cell(row=row_i, column=col_i).border = thin_border; ws.cell(row=row_i, column=col_i).alignment = center_align
        ws.cell(row=r+2, column=1+c_off, value="이름").font = Font(bold=True); ws.cell(row=r+2, column=2+c_off, value="월작정액").font = Font(bold=True)
        for i, m in enumerate(["01", "02", "03", "04", "05", "06"], 3): ws.cell(row=r+2, column=i+c_off, value=m).font = Font(bold=True)
        for i, m in enumerate(["07", "08", "09", "10", "11", "12"], 3): ws.cell(row=r+5, column=i+c_off, value=m).font = Font(bold=True)
        ws.merge_cells(start_row=r+3, start_column=1+c_off, end_row=r+7, end_column=1+c_off); ws.cell(row=r+3, column=1+c_off, value=f"{user['name']}\n{user['pos']}").font = Font(bold=True)
        ws.merge_cells(start_row=r+3, start_column=2+c_off, end_row=r+7, end_column=2+c_off); cc = ws.cell(row=r+3, column=2+c_off, value=user['commit'] if user['commit'] > 0 else "-")
        if user['commit'] > 0: cc.number_format = '#,##0'
        for i in range(6):
            ws.cell(row=r+3, column=i+3+c_off, value=user['labs'][i]); c1 = ws.cell(row=r+4, column=i+3+c_off, value=user['alloc'][i] if user['alloc'][i] > 0 else "")
            if user['alloc'][i] > 0: c1.number_format = '#,##0'
            ws.cell(row=r+6, column=i+3+c_off, value=user['labs'][i+6]); c2 = ws.cell(row=r+7, column=i+3+c_off, value=user['alloc'][i+6] if user['alloc'][i+6] > 0 else "")
            if user['alloc'][i+6] > 0: c2.number_format = '#,##0'
        ws.merge_cells(start_row=r+8, start_column=1+c_off, end_row=r+8, end_column=8+c_off); ws.cell(row=r+8, column=1+c_off, value="선교헌금에 관심가져주셔서 감사합니다.").alignment = center_align

    for i in range(0, len(user_list), 2):
        user_left, user_right = user_list[i], user_list[i+1] if i+1 < len(user_list) else None
        for r_idx in range(current_row, current_row + 9):
            ws.row_dimensions[r_idx].height = 25; ws.cell(row=r_idx, column=9).border = Border(right=Side(style='thin', color='000000'))
        draw_user_block(current_row, 0, user_left) 
        if user_right: draw_user_block(current_row, 10, user_right)
        pair_index = i // 2
        if pair_index % 2 == 0:
            gap_start = current_row + 9
            for r_idx in range(gap_start, gap_start + 3): ws.row_dimensions[r_idx].height = 25; ws.cell(row=r_idx, column=9).border = Border(right=Side(style='thin', color='000000'))
            for col_i in range(1, 19): ws.cell(row=gap_start, column=col_i).border = Border(bottom=Side(style='dashed', color='888888'), right=ws.cell(row=gap_start, column=col_i).border.right)
            current_row += 12 
        else:
            gap_start = current_row + 9; ws.row_dimensions[gap_start].height = 25; ws.cell(row=gap_start, column=9).border = Border(right=Side(style='thin', color='000000'))
            for col_i in range(1, 19): ws.cell(row=gap_start, column=col_i).border = Border(bottom=Side(style='dashed', color='888888'), right=ws.cell(row=gap_start, column=col_i).border.right)
            ws.row_breaks.append(Break(id=gap_start)); current_row += 10 
    ws.page_setup.orientation, ws.page_setup.fitToPage, ws.page_setup.fitToWidth = ws.ORIENTATION_LANDSCAPE, True, 1  
    ws.page_margins.left = ws.page_margins.right = 0.25
    output = io.BytesIO(); wb.save(output); return output.getvalue()

# --- 5. 앱 화면 구성 ---
for k in ['edit_idx_inc', 'edit_idx_exp', 'edit_idx_tgt', 'mode_inc', 'mode_exp', 'mode_tgt']:
    if k not in st.session_state: st.session_state[k] = None
with st.spinner('데이터 동기화 중...'):
    df_income, df_target, df_expense, raw_excel = load_data(FILE_ID)

if df_income is not None:
    if st.session_state["current_user"] in ["admin", "mission01"]:
        menu_options = ["🔍 개인별 조회", "✍️ 데이터 관리", "📊 결산/주단위집계", "🖨️ 인쇄용 집계표"]
    else: menu_options = ["📊 결산/주단위집계"]
    menu = st.sidebar.radio("메뉴", menu_options)
    
    t_n, t_p, t_a = get_col(df_target, ['이름', '성명'], 0), get_col(df_target, ['직분'], 1), get_col(df_target, ['월별 작정액', '작정액'], 2)
    i_n, i_y, i_a = get_col(df_income, ['이름', '성명'], 2), get_col(df_income, ['년월'], 1), get_col(df_income, ['금액'], 3)
    e_n, e_d, e_a = get_col(df_expense, ['내역'], 2), get_col(df_expense, ['날짜'], 0), get_col(df_expense, ['금액'], 3)

    if menu == "🔍 개인별 조회":
        names = [n for n in df_target[t_n].dropna().astype(str).str.strip().unique().tolist() if n and n != 'nan' and n != '합계']
        selected = st.selectbox("이름을 선택하세요", names)
        if selected:
            res = calculate_details(selected, df_income, df_target)
            if res:
                u_info = df_target[df_target[t_n].astype(str).str.strip() == selected]
                pos = clean_str(u_info.iloc[0].get(t_p, "")) if not u_info.empty else ""
                st.subheader(f"📄 {res['name']} ({pos})")
                st.write(f"기준일({datetime.now().strftime('%Y.%m.%d')}) 현재 / 월 작정액: {int(res['commit']):,}원 / 총 헌금액: {int(res['total']):,}원")
                html = "<table style='width:100%; border-collapse: collapse; text-align: center; margin-top: 15px; background-color: #ffffff; color: #333333;'>"
                html += "<tr style='background-color: #f8f9fa;'>"
                for i in range(1, 7): html += f"<th style='border: 1px solid #ddd; padding: 10px;'>{i}월</th>"
                html += "</tr><tr>"
                for i in range(6):
                    lab, amt = str(res['labs'][i]).replace('\n', '<br>'), f"{int(res['alloc'][i]):,}원" if res['alloc'][i] > 0 else "0원"
                    html += f"<td style='border: 1px solid #ddd; padding: 15px;'><span style='font-size:0.85em; color:#888;'>{lab}</span><br><b>{amt}</b></td>"
                html += "</tr><tr style='background-color: #f8f9fa;'>"
                for i in range(7, 13): html += f"<th style='border: 1px solid #ddd; padding: 10px;'>{i}월</th>"
                html += "</tr><tr>"
                for i in range(6, 12):
                    lab, amt = str(res['labs'][i]).replace('\n', '<br>'), f"{int(res['alloc'][i]):,}원" if res['alloc'][i] > 0 else "0원"
                    html += f"<td style='border: 1px solid #ddd; padding: 15px;'><span style='font-size:0.85em; color:#888;'>{lab}</span><br><b>{amt}</b></td>"
                html += "</tr></table>"; st.markdown(html, unsafe_allow_html=True)

    elif menu == "✍️ 데이터 관리":
        tab1, tab2, tab3 = st.tabs(["💰 헌금 수입", "📉 지출 내역", "👤 작정액 관리"])
        with tab1: 
            if st.session_state.mode_inc is None:
                # [최신순 정렬]
                df_view = df_income.copy()
                df_view['dt_sort'] = pd.to_datetime(df_view['날짜'], errors='coerce')
                df_view = df_view.sort_values(by='dt_sort', ascending=False).drop(columns=['dt_sort'])
                if '날짜' in df_view.columns: df_view['날짜'] = df_view['날짜'].apply(format_date_str)
                if i_a in df_view.columns: df_view[i_a] = pd.to_numeric(df_view[i_a], errors='coerce').fillna(0).apply(lambda x: f"{int(x):,} 원")
                disp = [c for c in df_view.columns if not str(c).startswith('Unnamed') and str(c) != i_y]
                # 표 높이 조절 (모바일 한 화면 가독성)
                st.dataframe(df_view[disp].dropna(subset=[i_n]), use_container_width=True, height=330)
                
                # [하단 고정 바 - 버튼 가로 한 줄 배치]
                with st.container():
                    st.markdown('<div class="fixed-footer">', unsafe_allow_html=True)
                    bc1, bc2, bc3, bc4 = st.columns([1.2, 1, 1, 1])
                    with bc1: 
                        if st.button("➕신규", key="inc_new", use_container_width=True): 
                            st.session_state.mode_inc = 'add'; st.rerun()
                    with bc2: 
                        idx = st.number_input("행", 0, len(df_income)-1, key="inc_idx_in", label_visibility="collapsed")
                    with bc3: 
                        if st.button("📝수정", key="inc_edit", use_container_width=True): 
                            st.session_state.edit_idx_inc, st.session_state.mode_inc = idx, 'edit'; st.rerun()
                    with bc4: 
                        if st.button("🗑️삭제", key="inc_del", use_container_width=True): 
                            st.session_state.edit_idx_inc, st.session_state.mode_inc = idx, 'delete_check'; st.rerun()
                    st.markdown('</div>', unsafe_allow_html=True)
                    
            elif st.session_state.mode_inc == 'add':
                with st.form("inc_add"):
                    d, amt = st.date_input("입금일자"), st.number_input("금액", min_value=0, step=1000)
                    opts = [f"{r[t_n]} ({r[t_p]})" if pd.notna(r.get(t_p)) else str(r.get(t_n)) for _, r in df_target.iterrows() if pd.notna(r.get(t_n)) and r.get(t_n)!='합계']
                    sel, note = st.selectbox("이름 선택", opts), st.text_input("비고")
                    if st.form_submit_button("저장"):
                        if save_to_drive(FILE_ID, append_dict_to_excel(raw_excel, '헌금수입', {'날짜': d.strftime("%Y-%m-%d"), i_y: d.strftime("%Y%m"), i_n: sel.split(" (")[0], i_a: amt, '비고': note})):
                            st.session_state.mode_inc = None; st.rerun()
                    if st.form_submit_button("취소"): st.session_state.mode_inc = None; st.rerun()
            elif st.session_state.mode_inc == 'edit':
                curr = df_income.iloc[st.session_state.edit_idx_inc]
                with st.form("inc_edit"):
                    new_d = st.date_input("날짜", value=pd.to_datetime(curr.get('날짜', datetime.now())) if pd.notna(curr.get('날짜')) else datetime.now())
                    new_n, new_a, new_b = st.text_input("이름", value=str(curr.get(i_n, ''))), st.number_input("금액", value=int(pd.to_numeric(curr.get(i_a, 0), errors='coerce') or 0), step=1000), st.text_input("비고", value=str(curr.get('비고', '')) if pd.notna(curr.get('비고')) else "")
                    if st.form_submit_button("✅ 수정 완료"):
                        df_income.loc[df_income.index[st.session_state.edit_idx_inc], ['날짜', i_y, i_n, i_a, '비고']] = [new_d.strftime("%Y-%m-%d"), new_d.strftime("%Y%m"), new_n, new_a, new_b]
                        if save_to_drive(FILE_ID, overwrite_sheet_preserve(raw_excel, '헌금수입', df_income)): st.session_state.mode_inc = None; st.rerun()
                    if st.form_submit_button("취소"): st.session_state.mode_inc = None; st.rerun()
            elif st.session_state.mode_inc == 'delete_check':
                st.warning(f"⚠️ {st.session_state.edit_idx_inc}번 행 데이터를 삭제하시겠습니까?")
                if st.button("🔴 삭제 실행", use_container_width=True):
                    df_income = df_income.drop(df_income.index[st.session_state.edit_idx_inc])
                    if save_to_drive(FILE_ID, overwrite_sheet_preserve(raw_excel, '헌금수입', df_income)): st.session_state.mode_inc = None; st.rerun()
                if st.button("취소", use_container_width=True): st.session_state.mode_inc = None; st.rerun()

        with tab2:
            if st.session_state.mode_exp is None:
                # [최신순 정렬]
                df_exp_v = df_expense.copy()
                df_exp_v['dt_sort'] = pd.to_datetime(df_exp_v['날짜'], errors='coerce')
                df_exp_v = df_exp_v.sort_values(by='dt_sort', ascending=False).drop(columns=['dt_sort'])
                if '날짜' in df_exp_v.columns: df_exp_v['날짜'] = df_exp_v['날짜'].apply(format_date_str)
                if '금액' in df_exp_v.columns: df_exp_v['금액'] = pd.to_numeric(df_exp_v['금액'], errors='coerce').fillna(0).apply(lambda x: f"{int(x):,} 원")
                st.dataframe(df_exp_v[[c for c in df_exp_v.columns if str(c) != '년월']], use_container_width=True, height=330)
                
                with st.container():
                    st.markdown('<div class="fixed-footer">', unsafe_allow_html=True)
                    ec1, ec2, ec3, ec4 = st.columns([1.2, 1, 1, 1])
                    with ec1: 
                        if st.button("➕지출", key="exp_new", use_container_width=True): st.session_state.mode_exp = 'add'; st.rerun()
                    with ec2: 
                        idx_e = st.number_input("행", 0, len(df_expense)-1, key="exp_idx_in", label_visibility="collapsed")
                    with ec3: 
                        if st.button("📝수정", key="exp_edit", use_container_width=True): st.session_state.edit_idx_exp, st.session_state.mode_exp = idx_e, 'edit'; st.rerun()
                    with ec4: 
                        if st.button("🗑️삭제", key="exp_del", use_container_width=True): st.session_state.edit_idx_exp, st.session_state.mode_exp = idx_e, 'delete_check'; st.rerun()
                    st.markdown('</div>', unsafe_allow_html=True)

            elif st.session_state.mode_exp == 'add':
                with st.form("exp_add"):
                    d, item, amt, note = st.date_input("지출일자"), st.text_input("지출항목"), st.number_input("금액", min_value=0, step=1000), st.text_input("비고")
                    if st.form_submit_button("저장"):
                        if save_to_drive(FILE_ID, append_dict_to_excel(raw_excel, '지출', {'날짜': d.strftime("%Y-%m-%d"), '년월': d.strftime("%Y%m"), '내역': item, '금액': amt, '비고': note})):
                            st.session_state.mode_exp = None; st.rerun()
                    if st.form_submit_button("취소"): st.session_state.mode_exp = None; st.rerun()
            elif st.session_state.mode_exp == 'edit':
                curr = df_expense.iloc[st.session_state.edit_idx_exp]
                with st.form("exp_edit"):
                    new_d = st.date_input("날짜", value=pd.to_datetime(curr.get('날짜', datetime.now())) if pd.notna(curr.get('날짜')) else datetime.now())
                    new_i, new_a, new_b = st.text_input("내역", value=str(curr.get('내역', ''))), st.number_input("금액", value=int(pd.to_numeric(curr.get('금액', 0), errors='coerce') or 0), step=1000), st.text_input("비고", value=str(curr.get('비고', '')) if pd.notna(curr.get('비고')) else "")
                    if st.form_submit_button("✅ 수정 완료"):
                        df_expense.loc[df_expense.index[st.session_state.edit_idx_exp], ['날짜', '년월', '내역', '금액', '비고']] = [new_d.strftime("%Y-%m-%d"), new_d.strftime("%Y%m"), new_i, new_a, new_b]
                        if save_to_drive(FILE_ID, overwrite_sheet_preserve(raw_excel, '지출', df_expense)): st.session_state.mode_exp = None; st.rerun()
                    if st.form_submit_button("취소"): st.session_state.mode_exp = None; st.rerun()
            elif st.session_state.mode_exp == 'delete_check':
                if st.button("🔴 지출 삭제 실행", use_container_width=True):
                    df_expense = df_expense.drop(df_expense.index[st.session_state.edit_idx_exp])
                    if save_to_drive(FILE_ID, overwrite_sheet_preserve(raw_excel, '지출', df_expense)): st.session_state.mode_exp = None; st.rerun()
                if st.button("취소", use_container_width=True): st.session_state.mode_exp = None; st.rerun()

        with tab3: 
            if st.session_state.mode_tgt is None:
                df_view = df_target.copy()
                for idx, row in df_view.iterrows():
                    name = clean_str(row.get(t_n))
                    if not name or name == 'nan' or name == '합계': continue
                    user_inc = df_income[df_income[i_n].apply(clean_str) == name]
                    total_donated = pd.to_numeric(user_inc[i_a], errors='coerce').sum()
                    m_amt = pd.to_numeric(row.get(t_a, 0), errors='coerce') or 0
                    df_view.loc[idx, ['년간작정금액', '헌금액', '년간작정 잔여금액']] = [m_amt * 12, total_donated, (m_amt * 12) - total_donated]
                for c in [t_a, '년간작정금액', '헌금액', '년간작정 잔여금액']: df_view[c] = pd.to_numeric(df_view[c], errors='coerce').fillna(0).apply(lambda x: f"{int(x):,} 원")
                disp = [c for c in df_view.columns if not str(c).startswith('Unnamed')]
                st.dataframe(df_view[disp], use_container_width=True, height=330)
                
                with st.container():
                    st.markdown('<div class="fixed-footer">', unsafe_allow_html=True)
                    tc1, tc2, tc3, tc4 = st.columns([1.2, 1, 1, 1])
                    with tc1: 
                        if st.button("➕신규", key="tgt_new", use_container_width=True): st.session_state.mode_tgt = 'add'; st.rerun()
                    with tc2: 
                        idx_t = st.number_input("행", 0, len(df_target)-1, key="tgt_idx_in", label_visibility="collapsed")
                    with tc3: 
                        if st.button("📝수정", key="tgt_edit", use_container_width=True): st.session_state.edit_idx_tgt, st.session_state.mode_tgt = idx_t, 'edit'; st.rerun()
                    with tc4: 
                        if st.button("🗑️삭제", key="tgt_del", use_container_width=True): st.session_state.edit_idx_tgt, st.session_state.mode_tgt = idx_t, 'delete_check'; st.rerun()
                    st.markdown('</div>', unsafe_allow_html=True)

            elif st.session_state.mode_tgt == 'add':
                with st.form("tgt_add"):
                    n, p, a = st.text_input("이름"), st.text_input("직분"), st.number_input("월별 작정액", min_value=0, step=1000)
                    if st.form_submit_button("저장"):
                        if save_to_drive(FILE_ID, append_dict_to_excel(raw_excel, '작정액', {t_n: n, t_p: p, t_a: a, '인쇄여부': 'N'})):
                            st.session_state.mode_tgt = None; st.rerun()
                    if st.form_submit_button("취소"): st.session_state.mode_tgt = None; st.rerun()
            elif st.session_state.mode_tgt == 'edit':
                curr = df_target.iloc[st.session_state.edit_idx_tgt]
                with st.form("tgt_edit"):
                    n, p, a = st.text_input("이름", value=str(curr.get(t_n, ''))), st.text_input("직분", value=str(curr.get(t_p, ''))), st.number_input("월별 작정액", value=int(pd.to_numeric(curr.get(t_a, 0), errors='coerce') or 0), step=1000)
                    if st.form_submit_button("✅ 수정 완료"):
                        df_target.loc[df_target.index[st.session_state.edit_idx_tgt], [t_n, t_p, t_a]] = [n, p, a]
                        if save_to_drive(FILE_ID, overwrite_sheet_preserve(raw_excel, '작정액', df_target)): st.session_state.mode_tgt = None; st.rerun()
                    if st.form_submit_button("취소"): st.session_state.mode_tgt = None; st.rerun()
            elif st.session_state.mode_tgt == 'delete_check':
                if st.button("🔴 성도 데이터 삭제 확정"):
                    df_target = df_target.drop(df_target.index[st.session_state.edit_idx_tgt])
                    if save_to_drive(FILE_ID, overwrite_sheet_preserve(raw_excel, '작정액', df_target)): st.session_state.mode_tgt = None; st.rerun()

    elif menu == "📊 결산/주단위집계":
        tab1, tab2 = st.tabs(["📅 월별 결산내역", "📆 주단위 결산내역"])
        df_inc_calc, df_exp_calc = df_income.copy(), df_expense.copy()
        df_inc_calc['amt'] = pd.to_numeric(df_inc_calc[i_a], errors='coerce').fillna(0)
        df_exp_calc['amt'] = pd.to_numeric(df_exp_calc[e_a], errors='coerce').fillna(0)
        c_inc = df_inc_calc[(df_inc_calc['날짜'].astype(str) < '2026-01-01') | (df_inc_calc[i_n].astype(str).str.contains('전년이월'))]['amt'].sum()
        c_exp = df_exp_calc[(df_exp_calc[e_d].astype(str) < '2026-01-01') | (df_exp_calc[e_n].astype(str).str.contains('전년이월'))]['amt'].sum()
        carryover_bal = c_inc - c_exp
        df_inc_26 = df_inc_calc[(df_inc_calc['날짜'].astype(str) >= '2026-01-01') & (~df_inc_calc[i_n].astype(str).str.contains('전년이월'))]
        df_exp_26 = df_exp_calc[(df_exp_calc[e_d].astype(str) >= '2026-01-01') & (~df_exp_calc[e_n].astype(str).str.contains('전년이월'))]
        with tab1:
            st.subheader("선교헌금 결산내역")
            monthly_data = [{"월별": "전년이월", "수입": carryover_bal, "지출": 0, "잔액": carryover_bal}]
            cur_bal, tot_inc, tot_exp = carryover_bal, carryover_bal, 0
            for m in range(1, 13):
                ym = f"2026{m:02d}"
                inc, exp = df_inc_26[df_inc_26[i_y] == ym]['amt'].sum(), df_exp_26[df_exp_26['년월'] == ym]['amt'].sum()
                if inc == 0 and exp == 0 and m > datetime.now().month: monthly_data.append({"월별": ym, "수입": 0, "지출": 0, "잔액": 0})
                else: cur_bal += (inc - exp); tot_inc += inc; tot_exp += exp; monthly_data.append({"월별": ym, "수입": inc, "지출": exp, "잔액": cur_bal})
            monthly_data.append({"월별": "합계", "수입": tot_inc, "지출": tot_exp, "잔액": tot_inc - tot_exp})
            h1 = "<table style='width:100%; border-collapse: collapse; text-align: center; border: 2px solid #a4b7c6; font-size: 15px; background-color: #ffffff; color: #333333;'>"
            h1 += "<tr style='background-color: #dbe5f1;'><th style='border: 1px solid #a4b7c6; padding: 10px;'>월별</th><th style='border: 1px solid #a4b7c6; padding: 10px;'>수입</th><th style='border: 1px solid #a4b7c6; padding: 10px;'>지출</th><th style='border: 1px solid #a4b7c6; padding: 10px;'>잔액</th></tr>"
            for row in monthly_data:
                bg = "#b4c6e7" if row['월별'] == "합계" else ("#f4f5f7" if row['월별'] == "전년이월" else "#ffffff")
                h1 += f"<tr style='background-color: {bg};'><td style='border: 1px solid #a4b7c6; padding: 8px;'>{row['월별']}</td><td style='border: 1px solid #a4b7c6; padding: 8px; text-align: right;'>{fmt(row['수입'])}</td><td style='border: 1px solid #a4b7c6; padding: 8px; text-align: right;'>{fmt(row['지출'])}</td><td style='border: 1px solid #a4b7c6; padding: 8px; text-align: right;'>{fmt(row['잔액'])}</td></tr>"
            st.markdown(h1 + "</table>", unsafe_allow_html=True)
        with tab2:
            st.subheader("선교헌금 주단위 결산내역")
            d_inc_list, d_exp_list = [format_date_str(d) for d in df_inc_26[df_inc_26['amt'] > 0]['날짜']], [format_date_str(d) for d in df_exp_26[df_exp_26['amt'] > 0][e_d]]
            all_dates = sorted(list(set([d for d in d_inc_list + d_exp_list if str(d).startswith('2026')])))
            weekly_temp = []
            cur_bal, tot_inc, tot_exp = carryover_bal, carryover_bal, 0
            for d_str in all_dates:
                inc, exp = df_inc_26[df_inc_26['날짜'].apply(format_date_str) == d_str]['amt'].sum(), df_exp_26[df_exp_26[e_d].apply(format_date_str) == d_str]['amt'].sum()
                cur_bal += (inc - exp); tot_inc += inc; tot_exp += exp; weekly_temp.append({"월별": d_str, "수입": inc, "지출": exp, "잔액": cur_bal})
            weekly_display = [{"월별": "합계", "수입": tot_inc, "지출": tot_exp, "잔액": tot_inc - tot_exp}] + weekly_temp[::-1] + [{"월별": "전년이월", "수입": carryover_bal, "지출": 0, "잔액": carryover_bal}]
            h2 = "<table style='width:100%; border-collapse: collapse; text-align: center; border: 2px solid #a4b7c6; font-size: 15px; background-color: #ffffff; color: #333333;'>"
            h2 += "<tr style='background-color: #dbe5f1;'><th style='border: 1px solid #a4b7c6; padding: 10px;'>월별</th><th style='border: 1px solid #a4b7c6; padding: 10px;'>수입</th><th style='border: 1px solid #a4b7c6; padding: 10px;'>지출</th><th style='border: 1px solid #a4b7c6; padding: 10px;'>잔액</th></tr>"
            for row in weekly_display:
                bg = "#b4c6e7" if row['월별'] == "합계" else ("#f4f5f7" if row['월별'] == "전년이월" else "#ffffff")
                h2 += f"<tr style='background-color: {bg};'><td style='border: 1px solid #a4b7c6; padding: 8px;'>{row['월별']}</td><td style='border: 1px solid #a4b7c6; padding: 8px; text-align: right;'>{fmt(row['수입'])}</td><td style='border: 1px solid #a4b7c6; padding: 8px; text-align: right;'>{fmt(row['지출'])}</td><td style='border: 1px solid #a4b7c6; padding: 8px; text-align: right;'>{fmt(row['잔액'])}</td></tr>"
            st.markdown(h2 + "</table>", unsafe_allow_html=True)

    elif menu == "🖨️ 인쇄용 집계표":
        st.subheader("🖨️ 인쇄용 엑셀 다운로드 (자동 가로 4명 출력)")
        months = sorted(list(set([f"2026{str(m).zfill(2)}" for m in range(1, 13)] + list(df_income[i_y].unique()))), reverse=True)
        target_month = st.selectbox("📌 기준월 선택", months)
        if st.button("🔄 인쇄 양식 엑셀 파일 만들기", use_container_width=True):
            with st.spinner("엑셀 2x2 그리드 배열을 그리는 중입니다..."):
                donors = set(df_income[df_income[i_y] == target_month][i_n].apply(clean_str).unique())
                for idx, row in df_target.iterrows():
                    name = clean_str(row.get(t_n))
                    if not name or name == 'nan' or name == '합계': continue
                    df_target.at[idx, '인쇄여부'] = 'Y' if name in donors else 'N'
                    user_inc = df_income[df_income[i_n].apply(clean_str) == name]
                    total_donated = pd.to_numeric(user_inc[i_a], errors='coerce').sum()
                    m_amt = pd.to_numeric(row.get(t_a, 0), errors='coerce') or 0
                    df_target.loc[idx, ['년간작정금액', '헌금액', '년간작정 잔여금액']] = [m_amt * 12, total_donated, (m_amt * 12) - total_donated]
                if save_to_drive(FILE_ID, overwrite_sheet_preserve(raw_excel, '작정액', df_target)):
                    st.session_state.download_data = generate_summary_excel(df_income, df_target, target_month); st.success(f"✅ 완성되었습니다!")
        if 'download_data' in st.session_state:
            st.download_button("📥 인쇄용 엑셀 다운로드", data=st.session_state.download_data, file_name=f"선교헌금_인쇄용_{target_month}.xlsx", use_container_width=True)
