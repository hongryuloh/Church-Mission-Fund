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
        # 이제 강제로 이름을 바꾸지 않고 엑셀 원본 그대로 가져옵니다. (안정성 극대화)
        df_income = pd.read_excel(io.BytesIO(raw_excel), sheet_name='헌금수입').astype(object)
        df_target = pd.read_excel(io.BytesIO(raw_excel), sheet_name='작정액').astype(object)
        try: df_expense = pd.read_excel(io.BytesIO(raw_excel), sheet_name='지출').astype(object)
        except: df_expense = pd.DataFrame(columns=['날짜', '년월', '내역', '금액', '비고']).astype(object)
            
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
    except Exception as e: return False

def overwrite_sheet(raw_excel, sheet_name, df_new):
    wb = openpyxl.load_workbook(io.BytesIO(raw_excel))
    if sheet_name in wb.sheetnames: del wb[sheet_name]
    ws = wb.create_sheet(sheet_name)
    ws.append(df_new.columns.tolist())
    for r in df_new.values.tolist(): 
        # NaN 값을 빈칸으로 안전하게 변환
        ws.append([None if pd.isna(val) else val for val in r])
    output = io.BytesIO(); wb.save(output); return output.getvalue()

# [핵심] 순서 무시하고 '이름'을 찾아 데이터를 정확히 넣는 스마트 저장 함수
def append_dict_to_excel(raw_excel, sheet_name, row_dict):
    wb = openpyxl.load_workbook(io.BytesIO(raw_excel))
    ws = wb[sheet_name] if sheet_name in wb.sheetnames else wb.create_sheet(sheet_name)
    
    header_map = {}
    for col_idx in range(1, ws.max_column + 1):
        val = ws.cell(row=1, column=col_idx).value
        if val is not None: header_map[str(val).strip()] = col_idx
            
    last_row = 1
    for row in range(ws.max_row, 0, -1):
        if any(ws.cell(row=row, column=c).value is not None for c in range(1, ws.max_column + 1)):
            last_row = row; break
    target_row = last_row + 1
    
    for key, val in row_dict.items():
        if key in header_map: ws.cell(row=target_row, column=header_map[key], value=val)
        else:
            new_col = ws.max_column + 1
            ws.cell(row=1, column=new_col, value=key)
            ws.cell(row=target_row, column=new_col, value=val)
            header_map[key] = new_col
            
    output = io.BytesIO(); wb.save(output); return output.getvalue()

def format_date_str(x):
    if pd.isna(x): return ""
    if isinstance(x, pd.Timestamp) or hasattr(x, 'strftime'): return x.strftime('%Y-%m-%d')
    x_str = str(x).strip().split(' ')[0].replace('.0', '')
    if len(x_str) == 8 and x_str.isdigit(): return f"{x_str[:4]}-{x_str[4:6]}-{x_str[6:8]}"
    if len(x_str) == 6 and x_str.isdigit(): return f"{x_str[:4]}-{x_str[4:6]}-01"
    return x_str

# --- 3. 데이터 계산 로직 (이름 기반으로 완벽 방어) ---
def calculate_details(user_name, df_income, df_target, start_year=2026):
    t_n = '성명' if '성명' in df_target.columns else df_target.columns[0]
    t_a = '작정액' if '작정액' in df_target.columns else df_target.columns[2]
    i_n = '성명' if '성명' in df_income.columns else df_income.columns[2]
    i_y = '년월' if '년월' in df_income.columns else df_income.columns[1]
    i_a = '금액' if '금액' in df_income.columns else df_income.columns[3]

    user_info = df_target[df_target[t_n].astype(str).str.strip() == user_name]
    if user_info.empty: return None
    commit = float(user_info.iloc[0][t_a]) if pd.notna(user_info.iloc[0][t_a]) else 0.0
    
    u_inc = df_income[df_income[i_n].astype(str).str.strip() == user_name].copy()
    u_inc['YM'] = u_inc[i_y].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
    paid = u_inc.groupby('YM').apply(lambda x: pd.to_numeric(x[i_a], errors='coerce').sum()).to_dict()
    
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

def generate_summary_excel(df_income, df_target, raw_excel, start_year=2026):
    wb = openpyxl.load_workbook(io.BytesIO(raw_excel))
    ws = wb['개인별 집계'] if '개인별 집계' in wb.sheetnames else wb.create_sheet('개인별 집계')
    ws.delete_rows(1, ws.max_row)
    ws.append(["성명", "작정액", "1월", "2월", "3월", "4월", "5월", "6월", "7월", "8월", "9월", "10월", "11월", "12월", "총납부액", "잔액"])
    t_n = '성명' if '성명' in df_target.columns else df_target.columns[0]
    names = [n for n in df_target[t_n].dropna().astype(str).str.strip().unique().tolist() if n and n.lower() != 'nan']
    for n in names:
        res = calculate_details(n, df_income, df_target, start_year)
        if res: ws.append([res["name"], res["commit"]] + res["alloc"] + [res["total"], max(0, res["commit"] - res["total"])])
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

    if menu == "🔍 개인별 조회":
        t_n = '성명' if '성명' in df_target.columns else df_target.columns[0]
        names = [n for n in df_target[t_n].dropna().astype(str).str.strip().unique().tolist() if n and n.lower() != 'nan']
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
        
        with tab1:
            if st.session_state.mode_inc is None:
                st.write("🔹 최근 헌금 수입")
                df_view = df_income.copy()
                if '날짜' in df_view.columns: df_view['날짜'] = df_view['날짜'].apply(format_date_str)
                if '금액' in df_view.columns: df_view['금액'] = pd.to_numeric(df_view['금액'], errors='coerce').fillna(0).apply(lambda x: f"{int(x):,} 원")
                if '성명' in df_view.columns: df_view = df_view.dropna(subset=['성명'])
                
                # Unnamed 빈 열들을 숨겨서 화면을 깔끔하게 유지합니다.
                disp_cols = [c for c in df_view.columns if not str(c).startswith('Unnamed')]
                st.dataframe(df_view[disp_cols], use_container_width=True)
                
                c1, c2, c3, c4 = st.columns([2,1,1,1])
                if c1.button("➕ 신규 등록", key="inc_a"): st.session_state.mode_inc = 'add'; st.rerun()
                idx = c2.number_input("행 번호", min_value=0, max_value=max(0, len(df_income)-1), step=1, key="inc_i")
                if c3.button("📝 수정", key="inc_e"): st.session_state.edit_idx_inc = idx; st.session_state.mode_inc = 'edit'; st.rerun()
                if c4.button("🗑️ 삭제", key="inc_d"): st.session_state.edit_idx_inc = idx; st.session_state.mode_inc = 'delete_check'; st.rerun()
            
            elif st.session_state.mode_inc == 'add':
                with st.form("inc_add"):
                    d = st.date_input("입금일자")
                    amt = st.number_input("금액", min_value=0, step=1000) # 1,000원 단위 적용
                    t_n = '성명' if '성명' in df_target.columns else df_target.columns[0]
                    t_p = '직분' if '직분' in df_target.columns else df_target.columns[1]
                    options = [f"{r[t_n]} ({r[t_p]})" if pd.notna(r.get(t_p)) else str(r.get(t_n)) for _, r in df_target.iterrows() if pd.notna(r.get(t_n))]
                    sel, note = st.selectbox("성명 선택", options), st.text_input("비고")
                    if st.form_submit_button("저장"):
                        # 위치(순서) 상관없이 지정된 이름의 열에 쏙쏙 들어갑니다.
                        row_dict = {'날짜': d.strftime("%Y-%m-%d"), '년월': d.strftime("%Y%m"), '성명': sel.split(" (")[0], '금액': amt, '비고': note}
                        if save_to_drive(FILE_ID, append_dict_to_excel(raw_excel, '헌금수입', row_dict)):
                            st.session_state.mode_inc = None; st.rerun()
                    if st.form_submit_button("취소"): st.session_state.mode_inc = None; st.rerun()

            elif st.session_state.mode_inc == 'delete_check':
                st.warning(f"⚠️ {st.session_state.edit_idx_inc}번 행을 삭제하시겠습니까?")
                if st.button("🔴 삭제 실행", key="inc_real_d"):
                    df_income = df_income.drop(df_income.index[st.session_state.edit_idx_inc])
                    if save_to_drive(FILE_ID, overwrite_sheet(raw_excel, '헌금수입', df_income)): st.session_state.mode_inc = None; st.rerun()
                if st.button("취소"): st.session_state.mode_inc = None; st.rerun()

            elif st.session_state.mode_inc == 'edit':
                curr = df_income.iloc[st.session_state.edit_idx_inc]
                with st.form("inc_edit"):
                    new_d = st.date_input("날짜", value=pd.to_datetime(curr.get('날짜', datetime.now())) if pd.notna(curr.get('날짜')) else datetime.now())
                    new_n = st.text_input("성명", value=str(curr.get('성명', '')))
                    new_a = st.number_input("금액", value=int(pd.to_numeric(curr.get('금액', 0), errors='coerce') or 0), step=1000)
                    new_b = st.text_input("비고", value=str(curr.get('비고', '')) if pd.notna(curr.get('비고')) else "")
                    
                    if st.form_submit_button("✅ 수정 완료"):
                        # 값 덮어쓰기 에러(ValueError)를 차단하는 .loc 방식
                        idx = df_income.index[st.session_state.edit_idx_inc]
                        df_income.loc[idx, '날짜'] = new_d.strftime("%Y-%m-%d")
                        df_income.loc[idx, '년월'] = new_d.strftime("%Y%m")
                        df_income.loc[idx, '성명'] = new_n
                        df_income.loc[idx, '금액'] = new_a
                        df_income.loc[idx, '비고'] = new_b
                        if save_to_drive(FILE_ID, overwrite_sheet(raw_excel, '헌금수입', df_income)):
                            st.session_state.mode_inc = None; st.rerun()
                    if st.form_submit_button("취소"): st.session_state.mode_inc = None; st.rerun()

        with tab2: # 지출
            if st.session_state.mode_exp is None:
                st.write("🔹 지출 내역")
                df_exp_view = df_expense.copy()
                if '날짜' in df_exp_view.columns: df_exp_view['날짜'] = df_exp_view['날짜'].apply(format_date_str)
                if '금액' in df_exp_view.columns: df_exp_view['금액'] = pd.to_numeric(df_exp_view['금액'], errors='coerce').fillna(0).apply(lambda x: f"{int(x):,} 원")
                disp_cols = [c for c in df_exp_view.columns if not str(c).startswith('Unnamed')]
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
                    if save_to_drive(FILE_ID, overwrite_sheet(raw_excel, '지출', df_expense)): st.session_state.mode_exp = None; st.rerun()
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
                        if save_to_drive(FILE_ID, overwrite_sheet(raw_excel, '지출', df_expense)):
                            st.session_state.mode_exp = None; st.rerun()
                    if st.form_submit_button("취소"): st.session_state.mode_exp = None; st.rerun()

        with tab3: # 작정액
            if st.session_state.mode_tgt is None:
                st.write("🔹 작정 명단")
                df_tgt_view = df_target.copy()
                if '작정액' in df_tgt_view.columns: df_tgt_view['작정액'] = pd.to_numeric(df_tgt_view['작정액'], errors='coerce').fillna(0).apply(lambda x: f"{int(x):,} 원")
                disp_cols = [c for c in df_tgt_view.columns if not str(c).startswith('Unnamed')]
                st.dataframe(df_tgt_view[disp_cols], use_container_width=True)
                
                c1, c2, c3, c4 = st.columns([2,1,1,1])
                if c1.button("➕ 신규 성도", key="tgt_a"): st.session_state.mode_tgt = 'add'; st.rerun()
                idx = c2.number_input("행 번호", min_value=0, max_value=max(0, len(df_target)-1), step=1, key="tgt_i")
                if c3.button("📝 수정", key="tgt_e"): st.session_state.edit_idx_tgt = idx; st.session_state.mode_tgt = 'edit'; st.rerun()
                if c4.button("🗑️ 삭제", key="tgt_d"): st.session_state.edit_idx_tgt = idx; st.session_state.mode_tgt = 'delete_check'; st.rerun()
            
            elif st.session_state.mode_tgt == 'add':
                with st.form("tgt_add"):
                    n, p = st.text_input("성명"), st.text_input("직분")
                    amt = st.number_input("월 작정액", min_value=0, step=1000)
                    if st.form_submit_button("저장"):
                        row_dict = {'성명': n, '직분': p, '작정액': amt}
                        if save_to_drive(FILE_ID, append_dict_to_excel(raw_excel, '작정액', row_dict)):
                            st.session_state.mode_tgt = None; st.rerun()
                    if st.form_submit_button("취소"): st.session_state.mode_tgt = None; st.rerun()

            elif st.session_state.mode_tgt == 'delete_check':
                st.warning("⚠️ 선택한 성도 정보를 삭제하시겠습니까?")
                if st.button("🔴 삭제 실행", key="tgt_real_d"):
                    df_target = df_target.drop(df_target.index[st.session_state.edit_idx_tgt])
                    if save_to_drive(FILE_ID, overwrite_sheet(raw_excel, '작정액', df_target)): st.session_state.mode_tgt = None; st.rerun()
                if st.button("취소"): st.session_state.mode_tgt = None; st.rerun()

            elif st.session_state.mode_tgt == 'edit':
                curr = df_target.iloc[st.session_state.edit_idx_tgt]
                with st.form("tgt_edit"):
                    n = st.text_input("성명", value=str(curr.get('성명', '')))
                    p = st.text_input("직분", value=str(curr.get('직분', '')))
                    a = st.number_input("작정액", value=int(pd.to_numeric(curr.get('작정액', 0), errors='coerce') or 0), step=1000)
                    
                    if st.form_submit_button("✅ 수정 완료"):
                        idx = df_target.index[st.session_state.edit_idx_tgt]
                        df_target.loc[idx, '성명'] = n
                        df_target.loc[idx, '직분'] = p
                        df_target.loc[idx, '작정액'] = a
                        if save_to_drive(FILE_ID, overwrite_sheet(raw_excel, '작정액', df_target)):
                            st.session_state.mode_tgt = None; st.rerun()
                    if st.form_submit_button("취소"): st.session_state.mode_tgt = None; st.rerun()

    elif menu == "📊 결산 및 통계":
        # 결산액도 '금액'이라는 이름표를 찾아서 정확히 더합니다.
        t_inc = pd.to_numeric(df_income['금액'], errors='coerce').sum() if '금액' in df_income.columns else 0
        t_exp = pd.to_numeric(df_expense['금액'], errors='coerce').sum() if '금액' in df_expense.columns else 0
        st.subheader("📊 전체 요약")
        c1, c2, c3 = st.columns(3)
        c1.metric("총 수입", f"{int(t_inc):,} 원"); c2.metric("총 지출", f"{int(t_exp):,} 원"); c3.metric("현재 잔액", f"{int(t_inc - t_exp):,} 원")

    elif menu == "🖨️ 인쇄용 집계표":
        with st.spinner("생성 중..."): data = generate_summary_excel(df_income, df_target, raw_excel)
        st.download_button("📥 엑셀 다운로드", data=data, file_name="최종집계.xlsx")
