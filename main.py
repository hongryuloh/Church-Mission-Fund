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

# --- 헬퍼 함수 ---
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

# --- 2. 구글 드라이브 데이터 로드 함수 ---
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

        def auto_fill_ym(df):
            if not df.empty and '날짜' in df.columns and '년월' in df.columns:
                for idx, row in df.iterrows():
                    val = str(row['년월']).strip()
                    if val in ['None', 'nan', '<NA>', 'NaT', '']:
                        d_val = format_date_str(row['날짜']).replace('-', '')
                        if len(d_val) >= 6: df.at[idx, '년월'] = d_val[:6]
            return df

        df_income = auto_fill_ym(robust_load('헌금수입', ['날짜', '년월', '이름', '금액', '비고']))
        df_target = robust_load('작정액', ['이름', '직분', '월별 작정액', '년간작정금액', '헌금액', '년간작정 잔여금액', '인쇄여부'])
        df_expense = auto_fill_ym(robust_load('지출', ['날짜', '년월', '내역', '금액', '비고']))
            
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
        
    header_map = {}
    for col_idx in range(1, ws.max_column + 1):
        val = ws.cell(row=header_row, column=col_idx).value
        if val is not None: header_map[str(val).strip()] = col_idx
            
    target_row = header_row + 1
    for _, row_data in df_new.iterrows():
        for col_name, val in row_data.items():
            col_str = str(col_name).strip()
            if col_str not in header_map and not col_str.startswith('Unnamed'):
                new_col = ws.max_column + 1
                ws.cell(row=header_row, column=new_col, value=col_str)
                header_map[col_str] = new_col
                
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
        if any(v in ['날짜', '이름', '성명', '내역', '작정액', '월별 작정액'] for v in vals):
            header_row = r; break
            
    header_map = {}
    for col_idx in range(1, ws.max_column + 1):
        val = ws.cell(row=header_row, column=col_idx).value
        if val is not None: header_map[str(val).strip()] = col_idx
            
    last_row = header_row
    for row in range(ws.max_row, header_row - 1, -1):
        if any(ws.cell(row=row, column=c).value is not None for c in range(1, ws.max_column + 1)):
            last_row = row; break
    target_row = last_row + 1
    
    for key, val in row_dict.items():
        if key not in header_map:
            new_col = ws.max_column + 1
            ws.cell(row=header_row, column=new_col, value=key)
            header_map[key] = new_col
        ws.cell(row=target_row, column=header_map[key], value=val)
            
    output = io.BytesIO(); wb.save(output); return output.getvalue()

# --- 3. 데이터 계산 로직 (사용자 피드백 전면 반영) ---
def calculate_details(user_name, df_income, df_target, start_year=2026):
    t_n = get_col(df_target, ['이름', '성명'], 0)
    t_a = get_col(df_target, ['월별 작정액', '작정액'], 2)
    i_n = get_col(df_income, ['이름', '성명'], 2)
    i_y = get_col(df_income, ['년월'], 1)
    i_a = get_col(df_income, ['금액'], 3)

    user_info = df_target[df_target[t_n].astype(str).str.strip() == user_name]
    if user_info.empty: return None
    
    # 1. 작정액 시트에서 이미 계산된 값(헌금액, 잔여금액)을 안전하게 바로 가져옵니다. (에러 원천 차단)
    commit = pd.to_numeric(user_info.iloc[0].get(t_a, 0), errors='coerce')
    commit = 0.0 if pd.isna(commit) else float(commit)
    
    total_donated = pd.to_numeric(user_info.iloc[0].get('헌금액', 0), errors='coerce')
    total_donated = 0.0 if pd.isna(total_donated) else float(total_donated)
    
    balance = pd.to_numeric(user_info.iloc[0].get('년간작정 잔여금액', 0), errors='coerce')
    balance = 0.0 if pd.isna(balance) else float(balance)

    # 2. 1~12월 상세 칸을 채우기 위해 헌금수입을 그룹화합니다.
    u_inc = df_income[df_income[i_n].astype(str).str.strip() == user_name].copy()
    u_inc['YM'] = u_inc[i_y].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
    u_inc[i_a] = pd.to_numeric(u_inc[i_a], errors='coerce').fillna(0) # 숫자 강제 변환으로 TypeError 방지
    paid = u_inc.groupby('YM')[i_a].sum().to_dict()
    
    alloc, lab = [0.0]*13, [""]*13
    sorted_m = sorted([m for m in paid.keys() if m.startswith(str(start_year))])
    
    if commit > 0:
        ptr = 1
        for pm in sorted_m:
            val = float(paid.get(pm, 0))
            while val > 0 and ptr <= 12:
                rem = commit - alloc[ptr]
                if rem <= 0: ptr += 1
                else:
                    amt = min(val, rem)
                    txt = f"{int(pm[-2:])}월납"
                    lab[ptr] = txt if lab[ptr] == "" else f"{lab[ptr]}<br>{txt}"
                    alloc[ptr] += amt; val -= amt
                    if alloc[ptr] >= commit - 0.0001: ptr += 1
    else:
        for pm in sorted_m:
            try:
                m = int(pm[-2:])
                if 1 <= m <= 12: alloc[m], lab[m] = float(paid.get(pm, 0)), f"{m}월납"
            except: pass 
            
    # sum 연산 삭제 완료 -> 엑셀에 있는 값(total_donated)을 그대로 반환!
    return {"name": user_name, "commit": commit, "alloc": alloc[1:], "labs": lab[1:], "total": total_donated, "balance": balance}

def generate_summary_excel(df_income, df_target, raw_excel, start_year=2026):
    wb = openpyxl.load_workbook(io.BytesIO(raw_excel))
    ws = wb['개인별 집계'] if '개인별 집계' in wb.sheetnames else wb.create_sheet('개인별 집계')
    ws.delete_rows(1, ws.max_row)
    
    # [수정] 작정액 시트의 정보를 인쇄용 집계표에도 일치시킵니다.
    headers = ["이름", "월별 작정액", "1월", "2월", "3월", "4월", "5월", "6월", "7월", "8월", "9월", "10월", "11월", "12월", "헌금액", "잔여금액", "인쇄여부"]
    ws.append(headers)
    
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    header_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    for cell in ws[1]:
        cell.font, cell.alignment, cell.border, cell.fill = Font(bold=True), Alignment(horizontal="center", vertical="center"), thin_border, header_fill

    t_n = get_col(df_target, ['이름', '성명'], 0)
    names = [n for n in df_target[t_n].dropna().astype(str).str.strip().unique().tolist() if n and n.lower() != 'nan']
    for n in names:
        res = calculate_details(n, df_income, df_target, start_year)
        
        # 작정액 시트에서 해당 성도의 '인쇄여부' 값을 그대로 가져옵니다.
        user_info = df_target[df_target[t_n].astype(str).str.strip() == n]
        print_status = user_info.iloc[0].get('인쇄여부', 'N') if not user_info.empty else 'N'
        
        if res: 
            row = [res["name"], res["commit"]] + res["alloc"] + [res["total"], res["balance"], print_status]
            ws.append(row)
            
            for col_idx, cell in enumerate(ws[ws.max_row], 1):
                cell.border = thin_border
                if col_idx == 1 or col_idx == ws.max_column: cell.alignment = Alignment(horizontal="center")
                else: cell.number_format = '#,##0'
                
    ws.column_dimensions['A'].width = 12
    for col in ['B', 'O', 'P', 'Q']: ws.column_dimensions[col].width = 14
    for c in ['C','D','E','F','G','H','I','J','K','L','M','N']: ws.column_dimensions[c].width = 11

    output = io.BytesIO(); wb.save(output); return output.getvalue()

# --- 5. 앱 화면 구성 ---
st.set_page_config(page_title="2026 선교헌금 관리", layout="wide")
st.title("⛪ 2026 선교헌금 관리 시스템")

for k in ['edit_idx_inc', 'edit_idx_exp', 'edit_idx_tgt', 'mode_inc', 'mode_exp', 'mode_tgt']:
    if k not in st.session_state: st.session_state[k] = None

with st.spinner('데이터 동기화 중...'):
    df_income, df_target, df_expense, raw_excel = load_data(FILE_ID)

if df_income is not None:
    menu = st.sidebar.radio("메뉴", ["🔍 개인별 조회", "✍️ 데이터 관리", "📊 결산 및 통계", "🖨️ 인쇄용 집계표"])

    # --- 공통 컬럼 별칭 (사용자 사진에 맞춰 '이름' 우선 검색) ---
    t_n = get_col(df_target, ['이름', '성명'], 0)
    t_p = get_col(df_target, ['직분'], 1)
    t_a = get_col(df_target, ['월별 작정액', '작정액'], 2)
    i_n = get_col(df_income, ['이름', '성명'], 2)
    i_y = get_col(df_income, ['년월'], 1)
    i_a = get_col(df_income, ['금액'], 3)

    if menu == "🔍 개인별 조회":
        names = [n for n in df_target[t_n].dropna().astype(str).str.strip().unique().tolist() if n and n.lower() != 'nan']
        selected = st.selectbox("성명을 선택하세요", names)
        if selected:
            res = calculate_details(selected, df_income, df_target)
            if res:
                st.subheader(f"📄 {res['name']} 성도님")
                st.write(f"월 작정액: {int(res['commit']):,}원 | 총 헌금액: {int(res['total']):,}원")
                for r in [0, 6]:
                    cols = st.columns(6)
                    for i in range(6):
                        m = r + i
                        with cols[i]:
                            st.write(f"**{m+1}월**"); st.markdown(f"<small>{res['labs'][m]}</small>", unsafe_allow_html=True); st.write(f"{int(res['alloc'][m]):,}원")

    elif menu == "✍️ 데이터 관리":
        tab1, tab2, tab3 = st.tabs(["💰 헌금 수입", "📉 지출 내역", "👤 작정액 관리"])
        
        with tab1: # 헌금 수입
            if st.session_state.mode_inc is None:
                st.write("🔹 최근 헌금 수입")
                df_view = df_income.copy()
                if '날짜' in df_view.columns: df_view['날짜'] = df_view['날짜'].apply(format_date_str)
                if i_a in df_view.columns: df_view[i_a] = pd.to_numeric(df_view[i_a], errors='coerce').fillna(0).apply(lambda x: f"{int(x):,} 원")
                if i_n in df_view.columns: df_view = df_view.dropna(subset=[i_n])
                disp_cols = [c for c in df_view.columns if not str(c).startswith('Unnamed') and str(c) != i_y]
                st.dataframe(df_view[disp_cols], use_container_width=True)
                
                c1, c2, c3, c4 = st.columns([2,1,1,1])
                if c1.button("➕ 신규 등록", key="inc_a"): st.session_state.mode_inc = 'add'; st.rerun()
                idx = c2.number_input("행 번호", min_value=0, max_value=max(0, len(df_income)-1), step=1, key="inc_i")
                if c3.button("📝 수정", key="inc_e"): st.session_state.edit_idx_inc = idx; st.session_state.mode_inc = 'edit'; st.rerun()
                if c4.button("🗑️ 삭제", key="inc_d"): st.session_state.edit_idx_inc = idx; st.session_state.mode_inc = 'delete_check'; st.rerun()
            
            elif st.session_state.mode_inc == 'add':
                with st.form("inc_add"):
                    d = st.date_input("입금일자")
                    amt = st.number_input("금액", min_value=0, step=1000) 
                    options = [f"{r[t_n]} ({r[t_p]})" if pd.notna(r.get(t_p)) else str(r.get(t_n)) for _, r in df_target.iterrows() if pd.notna(r.get(t_n))]
                    sel, note = st.selectbox("이름 선택", options), st.text_input("비고")
                    if st.form_submit_button("저장"):
                        row_dict = {'날짜': d.strftime("%Y-%m-%d"), i_y: d.strftime("%Y%m"), i_n: sel.split(" (")[0], i_a: amt, '비고': note}
                        if save_to_drive(FILE_ID, append_dict_to_excel(raw_excel, '헌금수입', row_dict)):
                            st.session_state.mode_inc = None; st.rerun()
                    if st.form_submit_button("취소"): st.session_state.mode_inc = None; st.rerun()

            elif st.session_state.mode_inc == 'delete_check':
                st.warning(f"⚠️ {st.session_state.edit_idx_inc}번 행을 삭제하시겠습니까?")
                if st.button("🔴 삭제 실행", key="inc_real_d"):
                    df_income = df_income.drop(df_income.index[st.session_state.edit_idx_inc])
                    if save_to_drive(FILE_ID, overwrite_sheet_preserve(raw_excel, '헌금수입', df_income)): st.session_state.mode_inc = None; st.rerun()
                if st.button("취소"): st.session_state.mode_inc = None; st.rerun()

            elif st.session_state.mode_inc == 'edit':
                curr = df_income.iloc[st.session_state.edit_idx_inc]
                with st.form("inc_edit"):
                    new_d = st.date_input("날짜", value=pd.to_datetime(curr.get('날짜', datetime.now())) if pd.notna(curr.get('날짜')) else datetime.now())
                    new_n = st.text_input("이름", value=str(curr.get(i_n, '')))
                    new_a = st.number_input("금액", value=int(pd.to_numeric(curr.get(i_a, 0), errors='coerce') or 0), step=1000)
                    new_b = st.text_input("비고", value=str(curr.get('비고', '')) if pd.notna(curr.get('비고')) else "")
                    
                    if st.form_submit_button("✅ 수정 완료"):
                        idx = df_income.index[st.session_state.edit_idx_inc]
                        df_income.loc[idx, '날짜'] = new_d.strftime("%Y-%m-%d")
                        df_income.loc[idx, i_y] = new_d.strftime("%Y%m")
                        df_income.loc[idx, i_n] = new_n
                        df_income.loc[idx, i_a] = new_a
                        df_income.loc[idx, '비고'] = new_b
                        if save_to_drive(FILE_ID, overwrite_sheet_preserve(raw_excel, '헌금수입', df_income)):
                            st.session_state.mode_inc = None; st.rerun()
                    if st.form_submit_button("취소"): st.session_state.mode_inc = None; st.rerun()

        with tab2: # 지출
            if st.session_state.mode_exp is None:
                st.write("🔹 지출 내역")
                df_exp_view = df_expense.copy()
                if '날짜' in df_exp_view.columns: df_exp_view['날짜'] = df_exp_view['날짜'].apply(format_date_str)
                if '금액' in df_exp_view.columns: df_exp_view['금액'] = pd.to_numeric(df_exp_view['금액'], errors='coerce').fillna(0).apply(lambda x: f"{int(x):,} 원")
                disp_cols = [c for c in df_exp_view.columns if not str(c).startswith('Unnamed') and str(c) != '년월']
                st.dataframe(df_exp_view[disp_cols], use_container_width=True)
                
                c1, c2, c3, c4 = st.columns([2,1,1,1])
                if c1.button("➕ 지출 등록", key="exp_a"): st.session_state.mode_exp = 'add'; st.rerun()
                idx = c2.number_input("행 번호", min_value=0, max_value=max(0, len(df_expense)-1), step=1, key="exp_i")
                if c3.button("📝 수정", key="exp_e"): st.session_state.edit_idx_exp = idx; st.session_state.mode_exp = 'edit'; st.rerun()
                if c4.button("🗑️ 삭제", key="exp_d"): st.session_state.edit_idx_exp = idx; st.session_state.mode_exp = 'delete_check'; st.rerun()
            
            elif st.session_state.mode_exp == 'add':
                with st.form("exp_add"):
                    d, item = st.date_input("지출일자"), st.text_input("지출항목")
                    amt = st.number_input("금액", min_value=0, step=1000)
                    note = st.text_input("비고")
                    if st.form_submit_button("저장"):
                        row_dict = {'날짜': d.strftime("%Y-%m-%d"), '년월': d.strftime("%Y%m"), '내역': item, '금액': amt, '비고': note}
                        if save_to_drive(FILE_ID, append_dict_to_excel(raw_excel, '지출', row_dict)):
                            st.session_state.mode_exp = None; st.rerun()
                    if st.form_submit_button("취소"): st.session_state.mode_exp = None; st.rerun()

            elif st.session_state.mode_exp == 'delete_check':
                st.warning(f"⚠️ 지출 {st.session_state.edit_idx_exp}번 삭제 확인")
                if st.button("🔴 삭제 실행", key="exp_real_d"):
                    df_expense = df_expense.drop(df_expense.index[st.session_state.edit_idx_exp])
                    if save_to_drive(FILE_ID, overwrite_sheet_preserve(raw_excel, '지출', df_expense)): st.session_state.mode_exp = None; st.rerun()
                if st.button("취소"): st.session_state.mode_exp = None; st.rerun()

            elif st.session_state.mode_exp == 'edit':
                curr = df_expense.iloc[st.session_state.edit_idx_exp]
                with st.form("exp_edit"):
                    new_d = st.date_input("날짜", value=pd.to_datetime(curr.get('날짜', datetime.now())) if pd.notna(curr.get('날짜')) else datetime.now())
                    new_i = st.text_input("내역", value=str(curr.get('내역', '')))
                    new_a = st.number_input("금액", value=int(pd.to_numeric(curr.get('금액', 0), errors='coerce') or 0), step=1000)
                    new_b = st.text_input("비고", value=str(curr.get('비고', '')) if pd.notna(curr.get('비고')) else "")
                    
                    if st.form_submit_button("✅ 수정 완료"):
                        idx = df_expense.index[st.session_state.edit_idx_exp]
                        df_expense.loc[idx, '날짜'] = new_d.strftime("%Y-%m-%d")
                        df_expense.loc[idx, '년월'] = new_d.strftime("%Y%m")
                        df_expense.loc[idx, '내역'] = new_i
                        df_expense.loc[idx, '금액'] = new_a
                        df_expense.loc[idx, '비고'] = new_b
                        if save_to_drive(FILE_ID, overwrite_sheet_preserve(raw_excel, '지출', df_expense)):
                            st.session_state.mode_exp = None; st.rerun()
                    if st.form_submit_button("취소"): st.session_state.mode_exp = None; st.rerun()

        with tab3: # 작정액 관리
            if st.session_state.mode_tgt is None:
                st.write("🔹 작정 명단 (데이터 자동 계산)")
                df_tgt_view = df_target.copy()
                
                for idx, row in df_tgt_view.iterrows():
                    name = str(row.get(t_n, '')).strip()
                    if not name or name == 'nan': continue
                    m_amt = pd.to_numeric(row.get(t_a, 0), errors='coerce')
                    m_amt = 0 if pd.isna(m_amt) else m_amt
                    y_amt = m_amt * 12
                    
                    user_donations = df_income[df_income[i_n].astype(str).str.strip() == name]
                    total_donated = pd.to_numeric(user_donations[i_a], errors='coerce').sum()
                    
                    df_tgt_view.loc[idx, '년간작정금액'] = y_amt
                    df_tgt_view.loc[idx, '헌금액'] = total_donated
                    df_tgt_view.loc[idx, '년간작정 잔여금액'] = y_amt - total_donated
                
                for col in [t_a, '년간작정금액', '헌금액', '년간작정 잔여금액']:
                    if col in df_tgt_view.columns:
                        df_tgt_view[col] = pd.to_numeric(df_tgt_view[col], errors='coerce').fillna(0).apply(lambda x: f"{int(x):,} 원")
                
                disp_cols = [c for c in df_tgt_view.columns if not str(c).startswith('Unnamed')]
                st.dataframe(df_tgt_view[disp_cols], use_container_width=True)
                
                c1, c2, c3, c4 = st.columns([2,1,1,1])
                if c1.button("➕ 신규 성도", key="tgt_a"): st.session_state.mode_tgt = 'add'; st.rerun()
                idx = c2.number_input("행 번호", min_value=0, max_value=max(0, len(df_target)-1), step=1, key="tgt_i")
                if c3.button("📝 수정", key="tgt_e"): st.session_state.edit_idx_tgt = idx; st.session_state.mode_tgt = 'edit'; st.rerun()
                if c4.button("🗑️ 삭제", key="tgt_d"): st.session_state.edit_idx_tgt = idx; st.session_state.mode_tgt = 'delete_check'; st.rerun()
            
            elif st.session_state.mode_tgt == 'add':
                with st.form("tgt_add"):
                    n, p = st.text_input("이름"), st.text_input("직분")
                    amt = st.number_input("월별 작정액", min_value=0, step=1000)
                    if st.form_submit_button("저장"):
                        user_donations = df_income[df_income[i_n].astype(str).str.strip() == n.strip()]
                        total_donated = pd.to_numeric(user_donations[i_a], errors='coerce').sum()
                        row_dict = {
                            t_n: n, t_p: p, t_a: amt, 
                            '년간작정금액': amt * 12, 
                            '헌금액': total_donated, 
                            '년간작정 잔여금액': (amt * 12) - total_donated,
                            '인쇄여부': 'N'
                        }
                        if save_to_drive(FILE_ID, append_dict_to_excel(raw_excel, '작정액', row_dict)):
                            st.session_state.mode_tgt = None; st.rerun()
                    if st.form_submit_button("취소"): st.session_state.mode_tgt = None; st.rerun()

            elif st.session_state.mode_tgt == 'delete_check':
                st.warning("⚠️ 선택한 성도 정보를 삭제하시겠습니까?")
                if st.button("🔴 삭제 실행", key="tgt_real_d"):
                    df_target = df_target.drop(df_target.index[st.session_state.edit_idx_tgt])
                    if save_to_drive(FILE_ID, overwrite_sheet_preserve(raw_excel, '작정액', df_target)): st.session_state.mode_tgt = None; st.rerun()
                if st.button("취소"): st.session_state.mode_tgt = None; st.rerun()

            elif st.session_state.mode_tgt == 'edit':
                curr = df_target.iloc[st.session_state.edit_idx_tgt]
                with st.form("tgt_edit"):
                    n = st.text_input("이름", value=str(curr.get(t_n, '')))
                    p = st.text_input("직분", value=str(curr.get(t_p, '')))
                    a = st.number_input("월별 작정액", value=int(pd.to_numeric(curr.get(t_a, 0), errors='coerce') or 0), step=1000)
                    
                    if st.form_submit_button("✅ 수정 완료"):
                        idx = df_target.index[st.session_state.edit_idx_tgt]
                        user_donations = df_income[df_income[i_n].astype(str).str.strip() == n.strip()]
                        total_donated = pd.to_numeric(user_donations[i_a], errors='coerce').sum()
                        
                        df_target.loc[idx, t_n] = n
                        df_target.loc[idx, t_p] = p
                        df_target.loc[idx, t_a] = a
                        df_target.loc[idx, '년간작정금액'] = a * 12
                        df_target.loc[idx, '헌금액'] = total_donated
                        df_target.loc[idx, '년간작정 잔여금액'] = (a * 12) - total_donated
                        
                        if save_to_drive(FILE_ID, overwrite_sheet_preserve(raw_excel, '작정액', df_target)):
                            st.session_state.mode_tgt = None; st.rerun()
                    if st.form_submit_button("취소"): st.session_state.mode_tgt = None; st.rerun()

    elif menu == "📊 결산 및 통계":
        t_inc = pd.to_numeric(df_income[i_a], errors='coerce').sum() if i_a in df_income.columns else 0
        t_exp = pd.to_numeric(df_expense['금액'], errors='coerce').sum() if '금액' in df_expense.columns else 0
        st.subheader("📊 전체 요약")
        c1, c2, c3 = st.columns(3)
        c1.metric("총 수입", f"{int(t_inc):,} 원"); c2.metric("총 지출", f"{int(t_exp):,} 원"); c3.metric("현재 잔액", f"{int(t_inc - t_exp):,} 원")

    elif menu == "🖨️ 인쇄용 집계표":
        st.subheader("🖨️ 개인별 집계표 생성 및 인쇄여부 일괄 업데이트")
        
        all_months = [f"2026{str(m).zfill(2)}" for m in range(1, 13)]
        data_months = [str(m) for m in df_income[i_y].unique() if str(m).isdigit() and len(str(m)) == 6]
        available_months = sorted(list(set(all_months + data_months)), reverse=True)
        
        target_month = st.selectbox("📌 기준월 선택 (해당 월에 헌금 내역이 있으면 '인쇄여부 Y'로 자동 업데이트됩니다)", available_months)
        
        if st.button("🔄 인쇄여부 업데이트 및 엑셀 다운로드 준비"):
            with st.spinner("작정액 시트 업데이트 및 엑셀 생성 중..."):
                for idx, row in df_target.iterrows():
                    name = str(row.get(t_n, '')).strip()
                    if not name or name == 'nan': continue
                    
                    # 1. 헌금수입 시트에서 선택한 월(target_month)의 해당 성도 헌금액 검사
                    donated_this_month = df_income[
                        (df_income[i_n].astype(str).str.strip() == name) & 
                        (df_income[i_y].astype(str).str.strip() == target_month)
                    ].copy()
                    
                    # 에러 차단을 위해 숫자 강제 변환
                    donated_this_month[i_a] = pd.to_numeric(donated_this_month[i_a], errors='coerce').fillna(0)
                    sum_this_month = donated_this_month[i_a].sum()
                    
                    # 2. 이번 달 헌금액이 있으면 Y, 없으면 N 반영
                    df_target.loc[idx, '인쇄여부'] = 'Y' if sum_this_month > 0 else 'N'
                    
                    # 3. 작정액 시트 값들도 엑셀 저장용으로 한 번 더 최신화
                    m_amt = pd.to_numeric(row.get(t_a, 0), errors='coerce')
                    m_amt = 0 if pd.isna(m_amt) else float(m_amt)
                    
                    user_donations = df_income[df_income[i_n].astype(str).str.strip() == name].copy()
                    user_donations[i_a] = pd.to_numeric(user_donations[i_a], errors='coerce').fillna(0)
                    total_donated = user_donations[i_a].sum()
                    
                    df_target.loc[idx, '년간작정금액'] = m_amt * 12
                    df_target.loc[idx, '헌금액'] = total_donated
                    df_target.loc[idx, '년간작정 잔여금액'] = (m_amt * 12) - total_donated

                # 구글 드라이브에 작정액 시트 덮어쓰기
                save_to_drive(FILE_ID, overwrite_sheet_preserve(raw_excel, '작정액', df_target))
                
                # 오류 100% 차단된 집계표 생성 함수 호출
                data = generate_summary_excel(df_income, df_target, raw_excel)
                st.session_state.download_data = data
                st.success(f"✅ {target_month}월 기준 '인쇄여부'가 업데이트되었습니다. 아래에서 엑셀을 다운로드하세요!")
                
        if 'download_data' in st.session_state:
            st.download_button("📥 엑셀 다운로드", data=st.session_state.download_data, file_name="최종집계.xlsx")
