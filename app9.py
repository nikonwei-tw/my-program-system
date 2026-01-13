import streamlit as st
import pandas as pd
import os
import io
import re

# --- 0. 頁面設定 ---
st.set_page_config(page_title="學分學程查詢與比對系統", layout="wide")

# --- 1. 自定義 CSS ---
st.markdown("""
    <style>
    .module-card {
        background-color: #f8f9fa; border-radius: 10px; padding: 18px;
        margin-bottom: 20px; border-left: 6px solid #007bff;
        box-shadow: 2px 2px 8px rgba(0,0,0,0.08);
    }
    .module-title { color: #0056b3; font-weight: bold; font-size: 1.15rem; margin-bottom: 5px; }
    .prog-container {
        background-color: #ffffff; border: 1px solid #e0e0e0;
        border-radius: 10px; padding: 20px; margin-bottom: 15px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
    }
    .module-summary-box {
        border-radius: 8px; padding: 12px 18px; margin: 10px 0;
        display: flex; justify-content: space-between; align-items: center;
        font-weight: bold;
    }
    .status-done { background-color: #d4edda; color: #155724; border-left: 6px solid #28a745; }
    .status-pending { background-color: #f1f3f5; color: #495057; border-left: 6px solid #adb5bd; }
    .note-box { background-color: #fff3cd; border-left: 5px solid #ffc107; padding: 12px; margin: 10px 0; font-size: 0.9rem; color: #856404; }
    </style>
    """, unsafe_allow_html=True)

# --- 2. 資料讀取 ---
@st.cache_data
def load_all_data(file_path):
    try:
        all_sheets = pd.read_excel(file_path, sheet_name=None, dtype=str)
        df_courses = all_sheets["科目表"]
        df_courses.columns = df_courses.columns.str.strip()
        
        # 核心：自動偵測認抵欄位名稱
        if '課程認抵範圍' not in df_courses.columns and '認抵單位代碼' in df_courses.columns:
            df_courses = df_courses.rename(columns={'認抵單位代碼': '課程認抵範圍'})
            
        df_courses['學分數'] = pd.to_numeric(df_courses['學分數'], errors='coerce').fillna(0)
        df_courses['模組名稱'] = df_courses.get('模組名稱', "未分類模組").fillna("未分類模組")
        
        prog_map = df_courses[['學程代碼', '學程名稱']].drop_duplicates().set_index('學程代碼')['學程名稱'].to_dict()
        df_summary = all_sheets.get("學程規範總額", pd.DataFrame())
        df_summary.columns = df_summary.columns.str.strip()
        
        if not df_summary.empty and '學程代碼' in df_summary.columns:
            df_summary['學程名稱'] = df_summary['學程代碼'].map(prog_map).fillna(df_summary.get('學程名稱', '未知學程'))
            for col in ['必修總學分', '選修總學分', '總計應修學分']:
                if col in df_summary.columns:
                    df_summary[col] = pd.to_numeric(df_summary[col], errors='coerce').fillna(0.0)
        
        return df_courses, df_summary, prog_map
    except Exception as e:
        st.error(f"讀取資料失敗，請檢查 Excel 欄位名稱。錯誤：{e}"); st.stop()

# --- 3. 核心工具 ---
def parse_required_credits(mod_name):
    match = re.search(r'\((\d+\.?\d*)\)', str(mod_name))
    return float(match.group(1)) if match else 0.0

def check_passing(grade):
    g = str(grade).strip().upper()
    if g in ["通過", "及格", "P", "PASS"]: return True
    try: return float(g) >= 60
    except: return False

def check_course_completion(req_row, passed_df):
    return not passed_df[passed_df['課程代碼'].str.strip() == str(req_row['課程代碼']).strip()].empty

def reset_filters():
    st.session_state.b1 = "全部"

# --- 4. 執行流程 ---
DATA_FILE = "master_data.xlsx"
if os.path.exists(DATA_FILE):
    df_courses, df_summary, prog_map = load_all_data(DATA_FILE)
else:
    st.error("找不到資料檔 master_data.xlsx"); st.stop()

tab_browse, tab_search, tab_audit = st.tabs(["📂 按學程瀏覽", "🔍 依課程反查學程", "🎓 學程完成度自動比對"])

# --- TAB 1: 瀏覽 ---
with tab_browse:
    if "b1" not in st.session_state: st.session_state.b1 = "全部"
    c1, c2, c3, c4 = st.columns([2, 2, 2, 1])
    with c1: sel_col = st.selectbox("1. 選擇學院", ["全部"] + sorted(df_courses["學院"].dropna().unique().tolist()), key="b1")
    with c2:
        temp_p = df_courses if sel_col == "全部" else df_courses[df_courses["學院"] == sel_col]
        sel_prog = st.selectbox("2. 選擇學程", sorted(temp_p["學程名稱"].dropna().unique().tolist()))
    with c3:
        year_list = sorted(df_courses[df_courses["學程名稱"]==sel_prog]["適用年度"].unique(), reverse=True)
        sel_year = st.selectbox("3. 選擇年度", year_list)
    with c4:
        st.markdown("<br>", unsafe_allow_html=True)
        st.button("🔄 重設", on_click=reset_filters)
    
    prog_data = df_courses[(df_courses["學程名稱"] == sel_prog) & (df_courses["適用年度"] == sel_year)]
    sum_row = df_summary[(df_summary["學程名稱"] == sel_prog) & (df_summary["適用年度"] == sel_year)]
    
    if not sum_row.empty:
        s = sum_row.iloc[0]
        st.success(f"**門檻：** 總計 {s.get('總計應修學分',0)} 學分")
    
    for cat in ["必修", "選修"]:
        cat_df = prog_data[prog_data["科目類別"] == cat]
        if not cat_df.empty:
            st.subheader(f"📍 {cat}課程")
            for mod in cat_df["模組名稱"].unique():
                mod_df = cat_df[cat_df["模組名稱"] == mod]
                st.markdown(f'<div class="module-card"><div class="module-title">🔹 {mod}</div>', unsafe_allow_html=True)
                # 使用 get 防止 KeyError
                display_cols = ['課程代碼', '課程名稱', '學分數', '課程認抵範圍', '備註']
                existing_cols = [c for c in display_cols if c in mod_df.columns]
                st.dataframe(mod_df[existing_cols], use_container_width=True, hide_index=True)
                st.markdown('</div>', unsafe_allow_html=True)

# --- TAB 2: 反查 ---
with tab_search:
    query = st.text_input("輸入課程關鍵字反查學程")
    if query:
        res = df_courses[df_courses["課程名稱"].str.contains(query, case=False, na=False)]
        st.dataframe(res[['學院', '學程名稱', '適用年度', '課程代碼', '課程名稱']].drop_duplicates(), use_container_width=True, hide_index=True)

# --- TAB 3: 比對 (結構化與互斥邏輯) ---
with tab_audit:
    st.header("🎓 學程達成度排行 (結構化計算)")
    uploaded_file = st.file_uploader("上傳成績單 Excel", type=["xlsx"])

    if uploaded_file:
        try:
            user_df = pd.read_excel(uploaded_file, dtype=str).dropna(how='all')
            user_df.columns = user_df.columns.str.strip()
            user_df['學分數'] = pd.to_numeric(user_df['學分數'], errors='coerce').fillna(0.0)
            user_df['及格'] = user_df['成績'].apply(check_passing)
            passed_df = user_df[user_df['及格']]

            audit_results = []
            for _, p_sum_row in df_summary.iterrows():
                p_courses = df_courses[(df_courses['學程名稱'] == p_sum_row['學程名稱']) & (df_courses['適用年度'] == p_sum_row['適用年度'])].copy()
                if p_courses.empty: continue
                
                # 互斥處理
                p_courses['已完成'] = p_courses.apply(lambda r: check_course_completion(r, passed_df), axis=1)
                p_courses['採計學分'] = p_courses['學分數']
                p_courses['互斥標記'] = ""
                used_mutex = set()
                for idx, row in p_courses.iterrows():
                    m_code = str(row.get('互斥代碼', '')).strip()
                    if row['已完成'] and m_code not in ['nan', '']:
                        if m_code in used_mutex:
                            p_courses.at[idx, '採計學分'] = 0.0
                            p_courses.at[idx, '互斥標記'] = "互斥不採計"
                        else:
                            used_mutex.add(m_code)

                # 結構化百分比計算
                mod_groups = p_courses.groupby('模組名稱')
                mod_scores = []
                for m_name, m_data in mod_groups:
                    m_req = parse_required_credits(m_name)
                    m_done = m_data.loc[m_data['已完成'], '採計學分'].sum()
                    mod_scores.append(min(m_done / m_req, 1.0) if m_req > 0 else (1.0 if m_done > 0 else 0.0))
                
                final_pct = sum(mod_scores) / len(mod_scores) if mod_scores else 0.0
                
                audit_results.append({
                    "name": p_sum_row['學程名稱'], "year": p_sum_row['適用年度'], 
                    "pct": final_pct, "details": p_courses, "goal": p_sum_row['總計應修學分']
                })

            audit_results.sort(key=lambda x: x['pct'], reverse=True)

            for res in audit_results:
                st.markdown('<div class="prog-container">', unsafe_allow_html=True)
                ca, cb = st.columns([4, 1])
                ca.markdown(f"#### 🎓 {res['name']} ({res['year']})")
                ca.progress(res['pct'])
                cb.markdown(f"### {int(res['pct']*100)}%")
                
                with st.expander("查看模組進度明細"):
                    for m_name in res['details']['模組名稱'].unique():
                        m_data = res['details'][res['details']['模組名稱'] == m_name]
                        m_done = m_data.loc[m_data['已完成'], '採計學分'].sum()
                        m_req = parse_required_credits(m_name)
                        is_sat = (m_done >= m_req) if m_req > 0 else (m_done > 0)
                        
                        st.markdown(f'<div class="module-summary-box {"status-done" if is_sat else "status-pending"}">'
                                    f'<span>{"✅" if is_sat else "⌛"} {m_name}</span>'
                                    f'<span>{int(m_done)}/{int(m_req)} 學分</span></div>', unsafe_allow_html=True)
                        
                        # 顯示要求的 7 個欄位 + 狀態
                        disp = m_data.copy()
                        disp['狀態'] = disp.apply(lambda r: "✅ 已達成" if r['已完成'] and r['採計學分']>0 else ("⚠️ 互斥不採計" if r['互斥標記'] else "❌ 未達成"), axis=1)
                        cols = ['科目類別', '課程代碼', '課程名稱', '學分數', '課程認抵範圍', '備註', '互斥代碼', '狀態']
                        st.table(disp[[c for c in cols if c in disp.columns]])
                st.markdown('</div>', unsafe_allow_html=True)
        except Exception as e:
            st.error(f"分析錯誤：{e}")