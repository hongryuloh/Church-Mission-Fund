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

# --- 1. 앱 기본 설정 및 로그인 로직 ---
st.set_page_config(page_title="2026 선교헌금 관리", layout="wide")

if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False
if "current_user" not in st.session_state:
    st.session_state["current_user"] = ""

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

# 상단 헤더
c1, c2 = st.columns([8, 1])
with c1: st.title("⛪ 2026 선교헌금 관리 시스템")
with c2: 
    st.write(f"👤 **{st.session_state['current_user']}**님")
    if st.button("로그아웃"):
        st.session_state["authenticated"] = False
        st.session_state["current_user"] = ""
        st.rerun()

# --- 2. 헬퍼 함수 및 데이터 연동 ---
FILE_ID = st.secrets["google"]["file_id"]

def get_gdrive_service():
    creds_json = json.loads(st.secrets["google"]["service_account"])
    credentials = service_account.Credentials.from_service_account_info(
        creds_json, scopes=['https://www.googleapis.com/auth/drive']
    )
    return build('drive', 'v3', credentials=credentials)

def clean_str(x): return str(x).split('.')[0].strip() if pd.notna(x) else ""

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

def fmt(val): return "-" if pd.isna(val) or val == 0 else f"{int(val):,}"

@st.cache_data(ttl=60) 
def load_data(file_id):
    try:
        service = get_gdrive_service()
        request = service.files().get_media(fileId=file_id)
        f_stream = io.BytesIO()
        downloader = MediaIoBaseDownload(f_stream, request)
        done = False
        while not done: _, done = downloader.next_chunk()
        raw_excel = f_stream.getvalue()
        
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

        df_inc = robust_load('헌금수입', ['날짜', '년월', '이름', '금액', '비고'])
        df_tgt = robust_load('작정액', ['이름', '직분', '월별 작정액', '년간작정금액', '헌금액', '년간작정 잔여금액', '인쇄여부'])
        df_exp = robust_load('지출', ['날짜', '년월', '내역', '금액', '비고'])
        
        for df in [df_inc, df_exp]:
            ym_col = get_col(df, ['년월'], 1)
            if ym_col:
                df[ym_col] = df[ym_col].apply(clean_str)
                for idx, row in df.iterrows():
                    if str(row[ym_col]).strip() in ['None', 'nan', '']:
                        d_val = format_date_str(row.get('날짜', '')).replace('-', '')
                        if len(d_val) >= 6: df.at[idx, ym_col] = d_val[:6]
        return df_inc, df_tgt, df_exp, raw_excel
    except Exception as e:
        st.error(f"데이터 로드 오류: {e}")
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
        if any(v in ['날짜', '이름', '성명', '내역', '작정액', '월별 작정액'] for v in vals): header_row = r; break
    if ws.max_row > header_row: ws.delete_rows(header_row + 1, ws.max_row - header_row)
    h_map = {str(ws.cell(row=header_row, column=c).value).strip(): c for c in range(1, ws.max_column + 1) if ws.cell(row=header_row, column=c).value is not None}
    t_row = header_row + 1
    for _, row_data in df_new.iterrows():
        for col_name, val in row_data.items():
            cs = str(col_name).strip()
            if cs in h_map: ws.cell(row=t_row, column=h_map[cs], value=None if pd.isna(val) else val)
        t_row += 1
    output = io.BytesIO(); wb.save(output); return output.getvalue()

def append_dict_to_excel(raw_excel, sheet_name, row_dict):
    wb = openpyxl.load_workbook(io.BytesIO(raw_excel))
    ws = wb[sheet_name] if sheet_name in wb.sheetnames else wb.create_sheet(sheet_name)
    header_row = 1
    for r in range(1, 11):
        vals = [str(ws.cell(row=r, column=c).value).strip() for c in range(1, ws.max_column + 1) if ws.cell(row=r, column=c).value is not None]
        if any(v in ['날짜', '이름', '성명', '내역', '작정액', '월별 작정액'] for v in vals): header_row = r; break
    h_map = {str(ws.cell(row=header_row, column=c).value).strip(): c for c in range(1, ws.max_column + 1) if ws.cell(row=header_row, column=c).value is not None}
    last_r = header_row
    for row in range(ws.max_row, header_row - 1, -1):
        if any(ws.cell(row=row, column=c).value is not None for c in range(1, ws.max_column + 1)): last_r = row; break
    t_row = last_r + 1
    for key, val in row_dict.items():
        if key not in h_map:
            new_c = ws.max_column + 1
            ws.cell(row=header_row, column=new_c, value=key); h_map[key] = new_c
        ws.cell(row=t_row, column=h_map[key], value=val)
    output = io.BytesIO(); wb.save(output); return output.getvalue()

# --- 3. 핵심 계산 로직 (작정액 기반 배분) ---
def calculate_details(user_name, df_income, df_target, start_year=2026):
    t_n, t_a = get_col(df_target, ['이름', '성명'], 0), get_col(df_target, ['월별 작정액', '작정액'], 2)
    i_n, i_y, i_a = get_col(df_income, ['이름', '성명'], 2), get_col(df_income, ['년월'], 1), get_col(df_income, ['금액'], 3)
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
                    amt = min(val, rem); alloc[ptr] += amt; val -= amt
                    txt = f"{int(str(pm)[-2:]):02d}월납"
                    lab[ptr] = txt if lab[ptr] == "" else f"{lab[ptr]}\n{txt}"
                    if alloc[ptr] >= commit - 0.01: ptr += 1
    else:
        for pm in sorted_m:
            try:
                m = int(str(pm)[-2:])
                if 1 <= m <= 12: 
                    alloc[m] = float(paid.get(pm, 0))
                    if alloc[m]>0: lab[m] = f"{m:02d}월납"
            except: pass
    return {"name": user_name, "commit": commit, "alloc": alloc[1:], "labs": lab[1:], "total": total}

# --- 4. 인쇄 포맷 (2x2 그리드 배열) ---
def generate_summary_excel(df_income, df_target, target_month, start_year=2026):
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "개인별 헌금내역"
    for c in range(1, 19): ws.column_dimensions[get_column_letter(c)].width = 8.13
    thin = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    t_n, t_p = get_col(df_target, ['이름', '성명'], 0), get_col(df_target, ['직분'], 1)
    user_list = []
    for _, row in df_target[df_target['인쇄여부'] == 'Y'].iterrows():
        name, pos = clean_str(row.get(t_n, '')), clean_str(row.get(t_p, ''))
        if not name or name in ['nan', '합계']: continue
        res = calculate_details(name, df_income, df_target, start_year)
        if res: res['pos'] = pos; user_list.append(res)
    
    cur_row = 1
    def draw_block(r, c_off, u):
        for ri in range(r, r+9): ws.row_dimensions[ri].height = 25
        ws.merge_cells(start_row=r, start_column=1+c_off, end_row=r, end_column=8+c_off)
        cell = ws.cell(row=r, column=1+c_off, value="2026년 선교헌금 작정 및 헌금내역")
        cell.font = Font(size=16, bold=True, underline="single"); cell.alignment = center
        ws.cell(row=r+1, column=4+c_off, value="선교헌금 합계 :").alignment = Alignment(horizontal='right')
        tc = ws.cell(row=r+1, column=7+c_off, value=u['total'])
        tc.font, tc.number_format = Font(bold=True), '#,##0'
        for ri in range(r+2, r+8):
            for ci in range(1+c_off, 9+c_off): ws.cell(row=ri, column=ci).border = thin; ws.cell(row=ri, column=ci).alignment = center
        ws.cell(row=r+3, column=1+c_off, value=f"{u['name']}\n{u['pos']}").font = Font(bold=True)
        ws.cell(row=r+3, column=2+c_off, value=u['commit']).number_format = '#,##0'
        for i in range(12):
            tr = r+3 if i < 6 else r+6
            ws.cell(row=tr, column=(i%6)+3+c_off, value=u['labs'][i])
            tr2 = r+4 if i < 6 else r+7
            amt_c = ws.cell(row=tr2, column=(i%6)+3+c_off, value=u['alloc'][i])
            if u['alloc'][i] > 0: amt_c.number_format = '#,##0'

    for i in range(0, len(user_list), 2):
        u_left, u_right = user_list[i], user_list[i+1] if i+1 < len(user_list) else None
        draw_block(cur_row, 0, u_left)
        if u_right: draw_block(cur_row, 10, u_right)
        cur_row += 12
        if (i // 2) % 2 == 1: ws.row_breaks.append(Break(id=cur_row-1))
    out = io.BytesIO(); wb.save(out); return out.getvalue()

# --- 5. 앱 화면 구성 (들여쓰기 및 기능 복구) ---
for k in ['edit_idx_inc', 'edit_idx_exp', 'edit_idx_tgt', 'mode_inc', 'mode_exp', 'mode_tgt']:
    if k not in st.session_state: st.session_state[k] = None

df_income, df_target, df_expense, raw_excel = load_data(FILE_ID)

if df_income is not None:
    # 권한 분리
    if st.session_state["current_user"] in ["admin", "mission01"]:
        menu_options = ["🔍 개인별 조회", "✍️ 데이터 관리", "📊 결산/주단위집계", "🖨️ 인쇄용 집계표"]
    else: menu_options = ["📊 결산/주단위집계"]
    
    menu = st.sidebar.radio("메뉴", menu_options)
    t_n, t_p, t_a = get_col(df_target, ['이름', '성명'], 0), get_col(df_target, ['직분'], 1), get_col(df_target, ['월별 작정액', '작정액'], 2)
    i_n, i_y, i_a = get_col(df_income, ['이름', '성명'], 2), get_col(df_income, ['년월'], 1), get_col(df_income, ['금액'], 3)
    e_n, e_d, e_a = get_col(df_expense, ['내역'], 2), get_col(df_expense, ['날짜'], 0), get_col(df_expense, ['금액'], 3)

    if menu == "🔍 개인별 조회":
        names = [n for n in df_target[t_n].dropna().unique() if n not in ['nan','합계']]
        sel = st.selectbox("이름을 선택하세요", names)
        if sel:
            res = calculate_details(sel, df_income, df_target)
            if res:
                st.subheader(f"📄 {sel} ({res.get('pos','')}) 성도님 현황")
                st.write(f"월 작정액: {int(res['commit']):,}원 / 총 헌금액: {int(res['total']):,}원")
                # 아이폰 다크모드 대응 표 (색상 고정)
                h = "<table style='width:100%; border-collapse: collapse; text-align: center; color: #333333; background-color: #ffffff;'>"
                h += "<tr style='background-color: #f8f9fa;'>" + "".join([f"<th style='border: 1px solid #ddd; padding: 10px;'>{i}월</th>" for i in range(1, 7)]) + "</tr><tr>"
                for i in range(6):
                    l, a = res['labs'][i].replace('\n','<br>'), f"{int(res['alloc'][i]):,}원"
                    h += f"<td style='border: 1px solid #ddd; padding: 10px;'><small style='color:#888;'>{l}</small><br><b>{a}</b></td>"
                h += "</tr><tr style='background-color: #f8f9fa;'>" + "".join([f"<th style='border: 1px solid #ddd; padding: 10px;'>{i}월</th>" for i in range(7, 13)]) + "</tr><tr>"
                for i in range(6, 12):
                    l, a = res['labs'][i].replace('\n','<br>'), f"{int(res['alloc'][i]):,}원"
                    h += f"<td style='border: 1px solid #ddd; padding: 10px;'><small style='color:#888;'>{l}</small><br><b>{a}</b></td>"
                h += "</tr></table>"
                st.markdown(h, unsafe_allow_html=True)

    elif menu == "✍️ 데이터 관리":
        tab1, tab2, tab3 = st.tabs(["💰 헌금 수입", "📉 지출 내역", "👤 작정액 관리"])
        with tab1: # 수입 관리 (복구 완료)
            if st.session_state.mode_inc is None:
                st.dataframe(df_income, use_container_width=True)
                c1, c2, c3, c4 = st.columns(4)
                if c1.button("➕ 수입 추가"): st.session_state.mode_inc = 'add'; st.rerun()
                idx = c2.number_input("행 번호", 0, len(df_income)-1, key="idx_inc")
                if c3.button("📝 수정"): st.session_state.edit_idx_inc, st.session_state.mode_inc = idx, 'edit'; st.rerun()
                if c4.button("🗑️ 삭제"): st.session_state.edit_idx_inc, st.session_state.mode_inc = idx, 'del'; st.rerun()
            elif st.session_state.mode_inc == 'add':
                with st.form("inc_a"):
                    d, n, a = st.date_input("날짜"), st.selectbox("이름", [x for x in df_target[t_n].unique() if x != '합계']), st.number_input("금액", step=1000)
                    if st.form_submit_button("저장"):
                        if save_to_drive(FILE_ID, append_dict_to_excel(raw_excel, '헌금수입', {'날짜':d.strftime("%Y-%m-%d"), '년월':d.strftime("%Y%m"), '이름':n, '금액':a, '비고':""})):
                            st.session_state.mode_inc = None; st.rerun()
            elif st.session_state.mode_inc == 'del':
                if st.button("🔴 정말 삭제하시겠습니까?"):
                    df_income = df_income.drop(df_income.index[st.session_state.edit_idx_inc])
                    if save_to_drive(FILE_ID, overwrite_sheet_preserve(raw_excel, '헌금수입', df_income)): st.session_state.mode_inc = None; st.rerun()
                if st.button("취소"): st.session_state.mode_inc = None; st.rerun()

        with tab2: # 지출 관리 (복구 완료)
            if st.session_state.mode_exp is None:
                st.dataframe(df_expense, use_container_width=True)
                if st.button("➕ 지출 추가"): st.session_state.mode_exp = 'add'; st.rerun()
            elif st.session_state.mode_exp == 'add':
                with st.form("exp_a"):
                    d, it, a = st.date_input("날짜"), st.text_input("내역"), st.number_input("금액", step=1000)
                    if st.form_submit_button("저장"):
                        if save_to_drive(FILE_ID, append_dict_to_excel(raw_excel, '지출', {'날짜':d.strftime("%Y-%m-%d"), '년월':d.strftime("%Y%m"), '내역':it, '금액':a, '비고':""})):
                            st.session_state.mode_exp = None; st.rerun()

        with tab3: # 성도 관리 (복구 완료)
            st.dataframe(df_target, use_container_width=True)

    elif menu == "📊 결산/주단위집계":
        tab1, tab2 = st.tabs(["📅 월별 결산", "📆 주단위 결산"])
        df_inc_c, df_exp_c = df_income.copy(), df_expense.copy()
        df_inc_c['amt'] = pd.to_numeric(df_inc_c[i_a], errors='coerce').fillna(0)
        df_exp_c['amt'] = pd.to_numeric(df_exp_c[e_a], errors='coerce').fillna(0)
        with tab1:
            st.subheader("📊 선교헌금 결산")
            m_data = []
            bal = 0
            for m in range(1, 13):
                ym = f"2026{m:02d}"
                inc = df_inc_c[df_inc_c[i_y].astype(str) == ym]['amt'].sum()
                exp = df_exp_c[df_exp_c['년월'].astype(str) == ym]['amt'].sum()
                bal += (inc - exp); m_data.append({"월별": ym, "수입": inc, "지출": exp, "잔액": bal})
            h = "<table style='width:100%; border-collapse: collapse; text-align: center; color: #333333; background-color:white;'>"
            h += "<tr style='background-color: #dbe5f1;'><th>월별</th><th>수입</th><th>지출</th><th>잔액</th></tr>"
            for r in m_data:
                h += f"<tr><td style='border: 1px solid #ddd;'>{r['월별']}</td><td style='border: 1px solid #ddd;'>{fmt(r['수입'])}</td><td style='border: 1px solid #ddd;'>{fmt(r['지출'])}</td><td style='border: 1px solid #ddd;'>{fmt(r['잔액'])}</td></tr>"
            st.markdown(h + "</table>", unsafe_allow_html=True)

    elif menu == "🖨️ 인쇄용 집계표":
        st.subheader("🖨️ 인쇄용 엑셀 다운로드")
        target_m = st.selectbox("기준월 선택", sorted(df_income[i_y].unique(), reverse=True))
        if st.button("🔄 인쇄 파일 생성"):
            donors = set(df_income[df_income[i_y] == target_m][i_n].astype(str).unique())
            df_target['인쇄여부'] = df_target[t_n].apply(lambda x: 'Y' if str(x) in donors else 'N')
            if save_to_drive(FILE_ID, overwrite_sheet_preserve(raw_excel, '작정액', df_target)):
                st.session_state.download_data = generate_summary_excel(df_income, df_target, target_m)
                st.success("✅ 완료되었습니다.")
        if 'download_data' in st.session_state:
            st.download_button("📥 엑셀 다운로드", data=st.session_state.download_data, file_name=f"인쇄_{target_m}.xlsx")
