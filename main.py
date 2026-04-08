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
@st.cache_data(ttl=60) 
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

# [공통] 시트 전체를 데이터프레임 내용으로 덮어쓰는 함수 (수정/삭제용)
def overwrite_sheet(raw_excel, sheet_name, df_new):
    wb = openpyxl.load_workbook(io.BytesIO(raw_excel))
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    ws = wb.create_sheet(sheet_name)
    
    # 헤더 작성
    ws.append(df_new.columns.tolist())
    # 데이터 작성
    for r in df_new.values.tolist():
        ws.append(r)
        
    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

def append_row_to_excel(raw_excel, sheet_name, row_data):
    wb = openpyxl.load_workbook(io.BytesIO(raw_excel))
    ws = wb[sheet_name] if sheet_name in wb.sheetnames else wb.create_sheet(sheet_name)
    
    # 실제 데이터가 있는 마지막 행 찾기
    last_row = 1
    for row in range(ws.max_row, 0, -1):
        if ws.cell(row=row, column=3).value is not None:
            last_row = row + 1
            break
    
    for col, val in enumerate(row_data, 1):
        ws.cell(row=last_row, column=col, value=val)
        
    output = io.BytesIO(); wb.save(output); return output.getvalue()

def format_date_str(x):
    if pd.isna(x): return ""
    if isinstance(x, pd.Timestamp) or hasattr(x, 'strftime'): return x.strftime('%Y-%m-%d')
    x_str = str(x).strip().split(' ')[0].replace('.0', '')
    if len(x_str) == 8 and x_str.isdigit(): return f"{x_str[:4]}-{x_str[4:6]}-{x_str[6:8]}"
    if len(x_str) == 6 and x_str.isdigit(): return f"{x_str[:4]}-{x_str[4:6]}-01"
    return x_str

# --- 3. 데이터 계산 로직 ---
def calculate_details(user_name, df_income, df_target, start_year=2026):
    user_info = df_target[df_target.iloc[:, 0].astype(str).str.strip() == user_name]
    if user_info.empty: return None
    commit = float(user_info.iloc[0, 2]) if pd.notna(user_info.iloc[0, 2]) else 0.0
    u_inc = df_income[df_income.iloc[:, 2].astype(str).str.strip() == user_name].copy()
    u_inc['YM'] = u_inc.iloc[:, 1].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
    paid = u_inc.groupby('YM').apply(lambda x: pd.to_numeric(x.iloc[:, 3], errors='coerce').sum()).to_dict()
    
    alloc, lab = [0.0]*13, [""]*13
    sorted_m = sorted([m for m in paid.keys() if m.startswith(str(start_year))])
    if commit > 0:
        ptr = 1
        for pm in sorted_m:
            val = float(paid[pm])
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
                if 1 <= m <= 12: alloc[m], lab[m] = paid[pm], f"{m}월납"
            except: pass 
    return {"name": user_name, "commit": commit, "alloc": alloc[1:], "labs": lab[1:], "total": sum(paid.values())}

# --- 4. 집계표 생성 ---
def generate_summary_excel(df_income, df_target, raw_excel, start_year=2026):
    wb = openpyxl.load_workbook(io.BytesIO(raw_excel))
    ws = wb['개인별 집계'] if '개인별 집계' in wb.sheetnames else wb.create_sheet('개인별 집계')
    ws.delete_rows(1, ws.max_row)
    ws.append(["성명", "작정액", "1월", "2월", "3월", "4월", "5월", "6월", "7월", "8월", "9월", "10월", "11월", "12월", "총납부액", "잔액"])
    names = [n for n in df_target.iloc[1:, 0].dropna().astype(str).str.strip().unique().tolist() if n and n.lower() != 'nan']
    for n in names:
        res = calculate_details(n, df_income, df_target, start_year)
        if res: ws.append([res["name"], res["commit"]] + res["alloc"] + [res["total"], max(0, res["commit"] - res["total"])])
    output = io.BytesIO(); wb.save(output); return output.getvalue()

# --- 5. 앱 화면 구성 ---
st.set_page_config(page_title="2026 선교헌금 관리", layout="wide")
st.title("⛪ 2026 선교헌금 관리 시스템")

# 세션 상태 초기화
for k in ['edit_idx_inc', 'edit_idx_exp', 'edit_idx_tgt', 'mode_inc', 'mode_exp', 'mode_tgt']:
    if k not in st.session_state: st.session_state[k] = None

with st.spinner('데이터 동기화 중...'):
    df_income, df_target, df_expense, raw_excel = load_data(FILE_ID)

if df_income is not None:
    menu = st.sidebar.radio("메뉴", ["🔍 개인별 조회", "✍️ 데이터 관리", "📊 결산 및 통계", "🖨️ 인쇄용 집계표"])

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
                        m = r + i
                        with cols[i]:
                            st.write(f"**{m+1}월**"); st.markdown(f"<small>{res['labs'][m]}</small>", unsafe_allow_html=True); st.write(f"{int(res['alloc'][m]):,}원")

    elif menu == "✍️ 데이터 관리":
        tab1, tab2, tab3 = st.tabs(["💰 헌금 수입", "📉 지출 내역", "👤 작정액 관리"])
        
        # --- TAB 1: 헌금 수입 ---
        with tab1:
            if st.session_state.mode_inc is None:
                st.write("🔹 최근 헌금 수입 (행 번호를 선택하여 수정/삭제 가능)")
                df_view = df_income.copy()
                df_view.columns = ['날짜', '년월', '성명', '금액', '비고'] + list(df_view.columns)[5:]
                df_view['날짜'] = df_view['날짜'].apply(format_date_str)
                df_view['금액'] = pd.to_numeric(df_view['금액'], errors='coerce').fillna(0).apply(lambda x: f"{int(x):,} 원")
                st.dataframe(df_view.iloc[1:].dropna(subset=['성명']), use_container_width=True)
                
                c1, c2, c3 = st.columns([2,1,1])
                if c1.button("➕ 신규 등록"): st.session_state.mode_inc = 'add'; st.rerun()
                idx = c2.number_input("수정/삭제할 행 번호", min_value=1, max_value=len(df_income)-1, step=1)
                if c3.button("📝 수정/삭제 실행"): st.session_state.edit_idx_inc = idx; st.session_state.mode_inc = 'edit'; st.rerun()
            
            elif st.session_state.mode_inc == 'add':
                with st.form("inc_add"):
                    d, amt = st.date_input("입금일자"), st.number_input("금액", min_value=0, step=10000)
                    options = [f"{r[0]} ({r[1]})" if pd.notna(r[1]) else str(r[0]) for r in df_target.iloc[1:, 0:2].values if pd.notna(r[0])]
                    sel, note = st.selectbox("성명 선택", options), st.text_input("비고")
                    if st.form_submit_button("저장"):
                        new = ["", d.strftime("%Y-%m-%d"), d.strftime("%Y%m"), sel.split(" (")[0], amt, note]
                        if save_to_drive(FILE_ID, append_row_to_excel(raw_excel, '헌금수입', new)):
                            st.session_state.mode_inc = None; st.rerun()
                    if st.form_submit_button("취소"): st.session_state.mode_inc = None; st.rerun()

            elif st.session_state.mode_inc == 'edit':
                curr = df_income.iloc[st.session_state.edit_idx_inc]
                with st.form("inc_edit"):
                    st.write(f"⚠️ {st.session_state.edit_idx_inc}번 행 수정 중")
                    new_d = st.date_input("날짜", value=pd.to_datetime(curr.iloc[1]) if pd.notna(curr.iloc[1]) else datetime.now())
                    new_n = st.text_input("성명", value=str(curr.iloc[2]))
                    new_a = st.number_input("금액", value=int(pd.to_numeric(curr.iloc[3], errors='coerce') or 0))
                    new_b = st.text_input("비고", value=str(curr.iloc[4]) if pd.notna(curr.iloc[4]) else "")
                    
                    if st.form_submit_button("✅ 수정 완료"):
                        df_income.iloc[st.session_state.edit_idx_inc, 1:5] = [new_d.strftime("%Y-%m-%d"), new_d.strftime("%Y%m"), new_n, new_a, new_b]
                        if save_to_drive(FILE_ID, overwrite_sheet(raw_excel, '헌금수입', df_income)):
                            st.session_state.mode_inc = None; st.rerun()
                    if st.form_submit_button("🗑️ 이 행 삭제", type="primary"):
                        df_income = df_income.drop(st.session_state.edit_idx_inc)
                        if save_to_drive(FILE_ID, overwrite_sheet(raw_excel, '헌금수입', df_income)):
                            st.session_state.mode_inc = None; st.rerun()
                    if st.form_submit_button("취소"): st.session_state.mode_inc = None; st.rerun()

        # --- TAB 2: 지출 내역 ---
        with tab2:
            if st.session_state.mode_exp is None:
                st.write("🔹 지출 내역")
                df_exp_view = df_expense.copy()
                df_exp_view['날짜'] = df_exp_view['날짜'].apply(format_date_str)
                df_exp_view['금액'] = pd.to_numeric(df_exp_view['금액'], errors='coerce').fillna(0).apply(lambda x: f"{int(x):,} 원")
                st.dataframe(df_exp_view, use_container_width=True)
                
                c1, c2, c3 = st.columns([2,1,1])
                if c1.button("➕ 지출 등록"): st.session_state.mode_exp = 'add'; st.rerun()
                idx = c2.number_input("수정/삭제 행 번호", min_value=0, max_value=len(df_expense)-1, step=1, key="exp_idx")
                if c3.button("📝 수정/삭제", key="exp_btn"): st.session_state.edit_idx_exp = idx; st.session_state.mode_exp = 'edit'; st.rerun()

            elif st.session_state.mode_exp == 'edit':
                curr = df_expense.iloc[st.session_state.edit_idx_exp]
                with st.form("exp_edit"):
                    st.write(f"⚠️ 지출 {st.session_state.edit_idx_exp}번 수정")
                    new_d = st.date_input("날짜", value=pd.to_datetime(curr['날짜']) if pd.notna(curr['날짜']) else datetime.now())
                    new_i = st.text_input("내역", value=str(curr['내역']))
                    new_a = st.number_input("금액", value=int(pd.to_numeric(curr['금액'], errors='coerce') or 0))
                    if st.form_submit_button("✅ 수정"):
                        df_expense.iloc[st.session_state.edit_idx_exp] = [new_d.strftime("%Y-%m-%d"), new_d.strftime("%Y%m"), new_i, new_a]
                        if save_to_drive(FILE_ID, overwrite_sheet(raw_excel, '지출', df_expense)):
                            st.session_state.mode_exp = None; st.rerun()
                    if st.form_submit_button("🗑️ 삭제", type="primary"):
                        df_expense = df_expense.drop(st.session_state.edit_idx_exp)
                        if save_to_drive(FILE_ID, overwrite_sheet(raw_excel, '지출', df_expense)):
                            st.session_state.mode_exp = None; st.rerun()
                    if st.form_submit_button("취소"): st.session_state.mode_exp = None; st.rerun()

        # --- TAB 3: 작정액 관리 ---
        with tab3:
            if st.session_state.mode_tgt is None:
                st.write("🔹 작정 명단")
                df_tgt_view = df_target.iloc[1:].copy()
                df_tgt_view.columns = ['성명', '직분', '작정액'] + list(df_tgt_view.columns)[3:]
                df_tgt_view['작정액'] = pd.to_numeric(df_tgt_view['작정액'], errors='coerce').fillna(0).apply(lambda x: f"{int(x):,} 원")
                st.dataframe(df_tgt_view, use_container_width=True)
                
                c1, c2, c3 = st.columns([2,1,1])
                if c1.button("➕ 신규 성도"): st.session_state.mode_tgt = 'add'; st.rerun()
                idx = c2.number_input("행 번호", min_value=1, max_value=len(df_target)-1, step=1, key="tgt_idx")
                if c3.button("📝 수정/삭제", key="tgt_btn"): st.session_state.edit_idx_tgt = idx; st.session_state.mode_tgt = 'edit'; st.rerun()

            elif st.session_state.mode_tgt == 'edit':
                curr = df_target.iloc[st.session_state.edit_idx_tgt]
                with st.form("tgt_edit"):
                    n, p, a = st.text_input("성명", value=str(curr.iloc[0])), st.text_input("직분", value=str(curr.iloc[1])), st.number_input("작정액", value=int(pd.to_numeric(curr.iloc[2], errors='coerce') or 0))
                    if st.form_submit_button("✅ 수정"):
                        df_target.iloc[st.session_state.edit_idx_tgt, 0:3] = [n, p, a]
                        if save_to_drive(FILE_ID, overwrite_sheet(raw_excel, '작정액', df_target)):
                            st.session_state.mode_tgt = None; st.rerun()
                    if st.form_submit_button("🗑️ 삭제", type="primary"):
                        df_target = df_target.drop(st.session_state.edit_idx_tgt)
                        if save_to_drive(FILE_ID, overwrite_sheet(raw_excel, '작정액', df_target)):
                            st.session_state.mode_tgt = None; st.rerun()
                    if st.form_submit_button("취소"): st.session_state.mode_tgt = None; st.rerun()

    elif menu == "📊 결산 및 통계":
        t_inc = pd.to_numeric(df_income.iloc[1:, 3], errors='coerce').sum()
        t_exp = pd.to_numeric(df_expense.iloc[:, 3], errors='coerce').sum() if not df_expense.empty else 0
        st.subheader("📊 전체 요약")
        c1, c2, c3 = st.columns(3)
        c1.metric("총 수입", f"{int(t_inc):,} 원"); c2.metric("총 지출", f"{int(t_exp):,} 원"); c3.metric("현재 잔액", f"{int(t_inc - t_exp):,} 원")

    elif menu == "🖨️ 인쇄용 집계표":
        with st.spinner("생성 중..."): data = generate_summary_excel(df_income, df_target, raw_excel)
        st.download_button("📥 엑셀 다운로드", data=data, file_name="최종집계.xlsx")
