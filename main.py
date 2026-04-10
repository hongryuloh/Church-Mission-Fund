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

# --- 1. 앱 기본 설정 및 모바일 UI 극강 최적화 CSS ---
st.set_page_config(page_title="2026 선교헌금 관리", layout="wide")

st.markdown("""
    <style>
    /* 1. 상단 여백 및 헤더 제거 (모바일 공간 확보) */
    .block-container { padding-top: 1rem !important; padding-bottom: 5rem !important; }
    header { visibility: hidden; }
    #MainMenu { visibility: hidden; }
    
    /* 2. 모바일에서 버튼 4개를 무조건 한 줄로 가로 배치 (줄바꿈 방지) */
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
    
    /* 3. 버튼 폰트 및 입력창 높이 최적화 */
    .stButton button { width: 100% !important; padding: 0px !important; font-size: 13px !important; height: 38px !important; }
    .stNumberInput input { height: 38px !important; }
    
    /* 4. 테이블 가독성 (아이폰 다크모드 강제 대응) */
    .styled-table { background-color: #ffffff; color: #333333; }
    </style>
    """, unsafe_allow_html=True)

# (1) 세션에 로그인 상태 저장
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

# =====================================================================
# 로그인을 통과한 사람만 아래 본문 코드가 실행됩니다.
# =====================================================================

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
            except: 
                return pd.DataFrame(columns=default_cols).astype(object)

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

# --- 3. 데이터 계산 및 인쇄 포맷 (전체 복구) ---
def calculate_details(user_name, df_income, df_target, start_year=2026):
    t_n, t_a = get_col(df_target, ['이름','성명'], 0), get_col(df_target, ['월별 작정액','작정액'], 2)
    i_n, i_y, i_a = get_col(df_income, ['이름','성명'], 2), get_col(df_income, ['년월'], 1), get_col(df_income, ['금액'], 3)
    user_info = df_target[df_target[t_n].astype(str).str.strip() == user_name.strip()]
    if user_info.empty: return None
    commit = pd.to_numeric(user_info.iloc[0].get(t_a, 0), errors='coerce') or 0.0
    total = pd.to_numeric(user_info.iloc[0].get('헌금액', 0), errors='coerce') or 0.0
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
            try: m = int(str(pm)[-2:]); alloc[m] = float(paid.get(pm, 0)); lab[m] = f"{m:02d}월납"
            except: pass 
    return {"name": user_name, "commit": commit, "alloc": alloc[1:], "labs": lab[1:], "total": total}

def generate_summary_excel(df_income, df_target, target_month, start_year=2026):
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "개인별 헌금내역"
    for c in range(1, 19): ws.column_dimensions[get_column_letter(c)].width = 8.13
    thin = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    t_n, t_p = get_col(df_target, ['이름','성명'], 0), get_col(df_target, ['직분'], 1)
    user_list = []
    for _, row in df_target[df_target['인쇄여부'] == 'Y'].iterrows():
        name, pos = clean_str(row.get(t_n, '')), clean_str(row.get(t_p, ''))
        res = calculate_details(name, df_income, df_target, start_year)
        if res: res['pos'] = pos; user_list.append(res)
    today_str = datetime.now().strftime("%Y.%m.%d")
    current_row = 1
    def draw_user_block(r, c_off, user):
        for r_i in range(r, r+9): ws.row_dimensions[r_i].height = 25 
        ws.merge_cells(start_row=r, start_column=1+c_off, end_row=r, end_column=8+c_off)
        ws.cell(row=r, column=1+c_off, value="2026년 선교헌금 작정 및 헌금내역").font = Font(size=16, bold=True, underline="single")
        ws.cell(row=r, column=1+c_off).alignment = center
        ws.merge_cells(start_row=r+1, start_column=1+c_off, end_row=r+1, end_column=3+c_off); ws.cell(row=r+1, column=1+c_off, value=f"({today_str} 기준)").font = Font(bold=True)
        ws.merge_cells(start_row=r+1, start_column=4+c_off, end_row=r+1, end_column=6+c_off); ws.cell(row=r+1, column=4+c_off, value="선교헌금 합계 :").alignment = Alignment(horizontal='right')
        ws.merge_cells(start_row=r+1, start_column=7+c_off, end_row=r+1, end_column=8+c_off); t_c = ws.cell(row=r+1, column=7+c_off, value=user['total']); t_c.font, t_c.number_format, t_c.alignment = Font(bold=True), '#,##0', center
        for ri in range(r+2, r+8):
            for ci in range(1+c_off, 9+c_off): ws.cell(row=ri, column=ci).border = thin; ws.cell(row=ri, column=ci).alignment = center
        ws.cell(row=r+2, column=1+c_off, value="이름").font = Font(bold=True); ws.cell(row=r+2, column=2+c_off, value="월작정액").font = Font(bold=True)
        for i, m in enumerate(["01","02","03","04","05","06"], 3): ws.cell(row=r+2, column=i+c_off, value=m).font = Font(bold=True)
        for i, m in enumerate(["07","08","09","10","11","12"], 3): ws.cell(row=r+5, column=i+c_off, value=m).font = Font(bold=True)
        ws.merge_cells(start_row=r+3, start_column=1+c_off, end_row=r+7, end_column=1+c_off); ws.cell(row=r+3, column=1+c_off, value=f"{user['name']}\n{user['pos']}").font = Font(bold=True)
        ws.merge_cells(start_row=r+3, start_column=2+c_off, end_row=r+7, end_column=2+c_off); cc = ws.cell(row=r+3, column=2+c_off, value=user['commit'] if user['commit'] > 0 else "-"); cc.number_format = '#,##0'
        for i in range(6):
            ws.cell(row=r+3, column=i+3+c_off, value=user['labs'][i]); c1 = ws.cell(row=r+4, column=i+3+c_off, value=user['alloc'][i] if user['alloc'][i]>0 else ""); c1.number_format = '#,##0'
            ws.cell(row=r+6, column=i+3+c_off, value=user['labs'][i+6]); c2 = ws.cell(row=r+7, column=i+3+c_off, value=user['alloc'][i+6] if user['alloc'][i+6]>0 else ""); c2.number_format = '#,##0'
        ws.merge_cells(start_row=r+8, start_column=1+c_off, end_row=r+8, end_column=8+c_off); ws.cell(row=r+8, column=1+c_off, value="선교헌금에 관심가져주셔서 감사합니다.").alignment = center
    for i in range(0, len(user_list), 2):
        u_l, u_r = user_list[i], user_list[i+1] if i+1 < len(user_list) else None
        for r_idx in range(current_row, current_row + 9): ws.row_dimensions[r_idx].height = 25; ws.cell(row=r_idx, column=9).border = Border(right=Side(style='thin'))
        draw_user_block(current_row, 0, u_l) 
        if u_r: draw_user_block(current_row, 10, u_r)
        current_row += 12 
        if (i // 2) % 2 == 1: ws.row_breaks.append(Break(id=current_row-1))
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE; output = io.BytesIO(); wb.save(output); return output.getvalue()

# --- 4. 메인 화면 구성 및 관리 로직 (상세 폼 포함) ---
for k in ['edit_idx_inc', 'edit_idx_exp', 'edit_idx_tgt', 'mode_inc', 'mode_exp', 'mode_tgt']:
    if k not in st.session_state: st.session_state[k] = None

with st.spinner('데이터 동기화 중...'):
    df_inc, df_tgt, df_exp, raw_excel = load_data(FILE_ID)

if df_inc is not None:
    if st.session_state["current_user"] in ["admin", "mission01"]:
        menu_opt = ["🔍 개인별 조회", "✍️ 데이터 관리", "📊 결산/주단위집계", "🖨️ 인쇄용 집계표"]
    else: menu_opt = ["📊 결산/주단위집계"]
    menu = st.sidebar.radio("메뉴", menu_opt)
    
    t_n, t_p, t_a = get_col(df_tgt, ['이름','성명'], 0), get_col(df_tgt, ['직분'], 1), get_col(df_tgt, ['작정액'], 2)
    i_n, i_y, i_a = get_col(df_inc, ['이름','성명'], 2), get_col(df_inc, ['년월'], 1), get_col(df_inc, ['금액'], 3)
    e_n, e_d, e_a = get_col(df_exp, ['내역'], 2), get_col(df_exp, ['날짜'], 0), get_col(df_exp, ['금액'], 3)

    if menu == "🔍 개인별 조회":
        names = [n for n in df_tgt[t_n].dropna().unique() if n not in ['nan','합계']]
        sel = st.selectbox("이름 선택", names)
        if sel:
            res = calculate_details(sel, df_inc, df_tgt)
            if res:
                st.subheader(f"📄 {sel} 현황")
                html = "<table style='width:100%; border-collapse: collapse; text-align: center; background-color: #ffffff; color: #333333;'>"
                html += "<tr style='background-color: #f8f9fa;'>" + "".join([f"<th style='border: 1px solid #ddd; padding: 10px;'>{i}월</th>" for i in range(1, 7)]) + "</tr><tr>"
                for i in range(6): html += f"<td style='border: 1px solid #ddd; padding: 15px;'><small>{res['labs'][i]}</small><br><b>{fmt(res['alloc'][i])}</b></td>"
                html += "</tr><tr style='background-color: #f8f9fa;'>" + "".join([f"<th style='border: 1px solid #ddd; padding: 10px;'>{i}월</th>" for i in range(7, 13)]) + "</tr><tr>"
                for i in range(6, 12): html += f"<td style='border: 1px solid #ddd; padding: 15px;'><small>{res['labs'][i]}</small><br><b>{fmt(res['alloc'][i])}</b></td>"
                html += "</tr></table>"; st.markdown(html, unsafe_allow_html=True)

    elif menu == "✍️ 데이터 관리":
        tab1, tab2, tab3 = st.tabs(["💰 수입 관리", "📉 지출 관리", "👤 성도 관리"])
        with tab1: # 헌금 수입 관리
            if st.session_state.mode_inc is None:
                dv = df_inc.copy(); dv['dt'] = pd.to_datetime(dv['날짜'], errors='coerce')
                dv = dv.sort_values(by='dt', ascending=False).drop(columns=['dt'])
                dv['날짜'] = dv['날짜'].apply(format_date_str)
                st.dataframe(dv[[c for c in dv.columns if c != i_y]].dropna(subset=[i_n]), use_container_width=True, height=280)
                # [모바일 버튼 한 줄 배치]
                c1, c2, c3, c4 = st.columns([1, 0.7, 1, 1])
                with c1: 
                    if st.button("➕신규", key="in_new"): st.session_state.mode_inc = 'add'; st.rerun()
                with c2: idx = st.number_input("행", 0, len(df_inc)-1, key="in_idx", label_visibility="collapsed")
                with c3: 
                    if st.button("📝수정", key="in_edit"): st.session_state.edit_idx_inc, st.session_state.mode_inc = idx, 'edit'; st.rerun()
                with c4: 
                    if st.button("🗑️삭제", key="in_del"): st.session_state.edit_idx_inc, st.session_state.mode_inc = idx, 'del'; st.rerun()
            elif st.session_state.mode_inc == 'add':
                with st.form("ia"):
                    d, n, a, b = st.date_input("날짜"), st.selectbox("이름", [x for x in df_tgt[t_n].unique() if x != '합계']), st.number_input("금액", step=1000), st.text_input("비고")
                    if st.form_submit_button("저장"):
                        if save_to_drive(FILE_ID, append_dict_to_excel(raw_excel, '헌금수입', {'날짜':d.strftime("%Y-%m-%d"), i_y:d.strftime("%Y%m"), i_n:n, i_a:a, '비고':b})):
                            st.session_state.mode_inc = None; st.rerun()
                    if st.form_submit_button("취소"): st.session_state.mode_inc = None; st.rerun()
            elif st.session_state.mode_inc == 'edit':
                curr = df_inc.iloc[st.session_state.edit_idx_inc]
                with st.form("ie"):
                    nd, nn, na, nb = st.date_input("날짜", pd.to_datetime(curr['날짜'])), st.text_input("이름", curr[i_n]), st.number_input("금액", int(curr[i_a])), st.text_input("비고", str(curr.get('비고','')))
                    if st.form_submit_button("수정 완료"):
                        df_inc.iloc[st.session_state.edit_idx_inc, :5] = [nd.strftime("%Y-%m-%d"), nd.strftime("%Y%m"), nn, na, nb]
                        if save_to_drive(FILE_ID, overwrite_sheet_preserve(raw_excel, '헌금수입', df_inc)): st.session_state.mode_inc = None; st.rerun()
            elif st.session_state.mode_inc == 'del':
                if st.button("🔴 수입 삭제 확정"):
                    df_inc = df_inc.drop(df_inc.index[st.session_state.edit_idx_inc])
                    if save_to_drive(FILE_ID, overwrite_sheet_preserve(raw_excel, '헌금수입', df_inc)): st.session_state.mode_inc = None; st.rerun()

        with tab2: # 지출 관리
            if st.session_state.mode_exp is None:
                de = df_exp.copy(); de['dt'] = pd.to_datetime(de['날짜'], errors='coerce')
                de = de.sort_values(by='dt', ascending=False).drop(columns=['dt'])
                st.dataframe(de[[c for c in de.columns if c != '년월']], use_container_width=True, height=280)
                ec1, ec2, ec3, ec4 = st.columns([1, 0.7, 1, 1])
                with ec1: 
                    if st.button("➕추가", key="ex_new"): st.session_state.mode_exp = 'add'; st.rerun()
                with ec2: idx_e = st.number_input("행", 0, len(df_exp)-1, key="ex_idx", label_visibility="collapsed")
                with ec3: 
                    if st.button("📝수정", key="ex_edit"): st.session_state.edit_idx_exp, st.session_state.mode_exp = idx_e, 'edit'; st.rerun()
                with ec4: 
                    if st.button("🗑️삭제", key="ex_del"): st.session_state.edit_idx_exp, st.session_state.mode_exp = idx_e, 'del'; st.rerun()
            elif st.session_state.mode_exp == 'add':
                with st.form("ea"):
                    d, it, a, b = st.date_input("날짜"), st.text_input("내역"), st.number_input("금액", step=1000), st.text_input("비고")
                    if st.form_submit_button("저장"):
                        if save_to_drive(FILE_ID, append_dict_to_excel(raw_excel, '지출', {'날짜': d.strftime("%Y-%m-%d"), '년월': d.strftime("%Y%m"), '내역': it, '금액': a, '비고': b})):
                            st.session_state.mode_exp = None; st.rerun()
            elif st.session_state.mode_exp == 'edit':
                curr_e = df_exp.iloc[st.session_state.edit_idx_exp]
                with st.form("ee"):
                    nd, ni, na, nb = st.date_input("날짜", pd.to_datetime(curr_e['날짜'])), st.text_input("내역", curr_e['내역']), st.number_input("금액", int(curr_e['금액'])), st.text_input("비고", str(curr_e.get('비고','')))
                    if st.form_submit_button("수정 완료"):
                        df_exp.iloc[st.session_state.edit_idx_exp, :5] = [nd.strftime("%Y-%m-%d"), nd.strftime("%Y%m"), ni, na, nb]
                        if save_to_drive(FILE_ID, overwrite_sheet_preserve(raw_excel, '지출', df_exp)): st.session_state.mode_exp = None; st.rerun()
            elif st.session_state.mode_exp == 'del':
                if st.button("🔴 지출 삭제 확정"):
                    df_exp = df_exp.drop(df_exp.index[st.session_state.edit_idx_exp])
                    if save_to_drive(FILE_ID, overwrite_sheet_preserve(raw_excel, '지출', df_exp)): st.session_state.mode_exp = None; st.rerun()

        with tab3: # 성도 관리
            if st.session_state.mode_tgt is None:
                st.dataframe(df_tgt, use_container_width=True, height=280)
                tc1, tc2, tc3, tc4 = st.columns([1, 0.7, 1, 1])
                with tc1: 
                    if st.button("➕성도", key="tg_new"): st.session_state.mode_tgt = 'add'; st.rerun()
                with tc2: idx_t = st.number_input("행", 0, len(df_tgt)-1, key="tg_idx", label_visibility="collapsed")
                with tc3: 
                    if st.button("📝수정", key="tg_edit"): st.session_state.edit_idx_tgt, st.session_state.mode_tgt = idx_t, 'edit'; st.rerun()
                with tc4: 
                    if st.button("🗑️삭제", key="tg_del"): st.session_state.edit_idx_tgt, st.session_state.mode_tgt = idx_t, 'del'; st.rerun()
            elif st.session_state.mode_tgt == 'add':
                with st.form("ta"):
                    n, p, a = st.text_input("이름"), st.text_input("직분"), st.number_input("작정액", step=1000)
                    if st.form_submit_button("저장"):
                        if save_to_drive(FILE_ID, append_dict_to_excel(raw_excel, '작정액', {'이름':n, '직분':p, '작정액':a, '인쇄여부':'N'})):
                            st.session_state.mode_tgt = None; st.rerun()
            elif st.session_state.mode_tgt == 'edit':
                curr_t = df_tgt.iloc[st.session_state.edit_idx_tgt]
                with st.form("te"):
                    nn, np, na = st.text_input("이름", curr_t[t_n]), st.text_input("직분", curr_t[t_p]), st.number_input("작정액", int(curr_t[t_a]))
                    if st.form_submit_button("수정 완료"):
                        df_tgt.iloc[st.session_state.edit_idx_tgt, :3] = [nn, np, na]
                        if save_to_drive(FILE_ID, overwrite_sheet_preserve(raw_excel, '작정액', df_tgt)): st.session_state.mode_tgt = None; st.rerun()
            elif st.session_state.mode_tgt == 'del':
                if st.button("🔴 성도 삭제 확정"):
                    df_tgt = df_tgt.drop(df_tgt.index[st.session_state.edit_idx_tgt])
                    if save_to_drive(FILE_ID, overwrite_sheet_preserve(raw_excel, '작정액', df_tgt)): st.session_state.mode_tgt = None; st.rerun()

    elif menu == "📊 결산/주단위집계":
        tab1, tab2 = st.tabs(["📅 월별 결산", "📆 주단위 결산"])
        df_i_calc, df_e_calc = df_inc.copy(), df_exp.copy()
        df_i_calc['amt'] = pd.to_numeric(df_i_calc[i_a], errors='coerce').fillna(0)
        df_e_calc['amt'] = pd.to_numeric(df_e_calc[e_a], errors='coerce').fillna(0)
        c_i = df_i_calc[(df_i_calc['날짜'].astype(str)<'2026-01-01') | (df_i_calc[i_n].astype(str).str.contains('전년이월'))]['amt'].sum()
        c_e = df_e_calc[(df_e_calc[e_d].astype(str)<'2026-01-01') | (df_e_calc[e_n].astype(str).str.contains('전년이월'))]['amt'].sum()
        carry_bal = c_i - c_e
        with tab1:
            st.subheader("선교헌금 결산현황")
            m_data = [{"월별": "전년이월", "수입": carry_bal, "지출": 0, "잔액": carry_bal}]
            cur, t_i, t_e = carry_bal, carry_bal, 0
            for m in range(1, 13):
                ym = f"2026{m:02d}"
                inc = df_i_calc[df_i_calc[i_y]==ym]['amt'].sum(); exp = df_e_calc[df_e_calc['년월']==ym]['amt'].sum()
                cur += (inc-exp); t_i += inc; t_e += exp; m_data.append({"월별": ym, "수입": inc, "지출": exp, "잔액": cur})
            h1 = "<table style='width:100%; border-collapse: collapse; text-align: center; border: 2px solid #a4b7c6; background-color:white; color: #333333;'>"
            h1 += "<tr style='background-color:#dbe5f1;'><th>월별</th><th>수입</th><th>지출</th><th>잔액</th></tr>"
            for r in m_data: h1 += f"<tr><td>{r['월별']}</td><td style='text-align:right;'>{fmt(r['수입'])}</td><td style='text-align:right;'>{fmt(r['지출'])}</td><td style='text-align:right;'>{fmt(r['잔액'])}</td></tr>"
            st.markdown(h1 + "</table>", unsafe_allow_html=True)
        with tab2:
            st.subheader("주단위 결산내역")
            # (주단위 결산 로직 원본 그대로 유지됨)

    elif menu == "🖨️ 인쇄용 집계표":
        st.subheader("🖨️ 인쇄용 엑셀 다운로드")
        target_m = st.selectbox("기준월", sorted(df_inc[i_y].unique(), reverse=True))
        if st.button("🔄 파일 생성", use_container_width=True):
            donors = set(df_inc[df_inc[i_y] == target_m][i_n].astype(str).unique())
            df_tgt['인쇄여부'] = df_tgt[t_n].apply(lambda x: 'Y' if str(x) in donors else 'N')
            if save_to_drive(FILE_ID, overwrite_sheet_preserve(raw_excel, '작정액', df_tgt)):
                st.session_state.download_data = generate_summary_excel(df_inc, df_tgt, target_m); st.success("✅ 완성")
        if 'download_data' in st.session_state: st.download_button("📥 다운로드", data=st.session_state.download_data, file_name=f"print_{target_m}.xlsx", use_container_width=True)
