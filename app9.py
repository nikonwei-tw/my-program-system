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
    .status-done {
        background-color: #d4edda; color: #155724;
        border-left: 6px solid #28a745;
    }
    .status-pending {
        background-color: #f1f3f5; color: #495057;
        border-left: 6px solid #adb5bd;
    }
    .note-box {
        background-color: #fff3cd; border-left: 5px solid #ffc107;
        padding: 12px; margin: 10px 0; font-size: 0.9rem; color: #856404;
    }
    </style>
    """, unsafe_allow_html=True)

# --- 2. 資料讀取 ---
@st.cache_data
def load_all_data(file_path):
    try:
        all_sheets = pd.read_excel(file_path, sheet_name=None, dtype=str)
        df_courses = all_sheets["科目表"]
        df_courses.columns = df_courses.columns.str.strip()
        df_courses['學分數'] = pd.to_numeric(df_courses['學分數'], errors='coerce').fillna(0)
        
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
        st.error(f"讀取資料失敗：{e}"); st.stop()

# --- 3. 核心工具函數 ---
def parse_required_credits(mod_name):
    """提取模組名稱括號中的數字，如 '必修模組(8)' -> 8.0"""
    match = re.search(r'\((\d+\.?\d*)\)', mod_name)
    return float(match.group(1)) if match else 0.0

def check_passing(grade):
    g = str(grade).strip().upper()
    if g in ["通過", "及格", "P", "PASS"]: return True
    try: return float(g) >= 60
    except: return False

def check_course_completion(req_row, passed_df):
    matches = passed_df[passed_df['課程代碼'].str.strip() == str(req_row['課程代碼']).strip()]
    return not matches.empty

def reset_filters():
    st.session_state.b1 = "全部"

# --- 4. 執行流程 ---
DATA_FILE = "master_data.xlsx"
if os.path.exists(DATA_FILE):
    df_courses, df_summary, prog_map = load_all_data(DATA_FILE)
else:
    st.error("找不到資料檔 master_data.xlsx"); st.stop()

tab_browse, tab_search, tab_audit = st.tabs(["📂 按學程瀏覽", "🔍 依課程反查學程", "🎓 學程完成度自動比對"])

# --- TAB 1: 按學程瀏覽 ---
with tab_browse:
    if "b1" not in st.session_state: st.session_state.b1 = "全部"
    c1, c2, c3, c4 = st.columns([2, 2, 2, 1])
    with c1: 
        sel_col = st.selectbox("1. 選擇學院", ["全部"] + sorted(df_courses["學院"].dropna().unique().tolist()), key="b1")
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
        st.success(f"**門檻要求：** 必修 {s.get('必修總學分',0)} / 選修 {s.get('選修總學分',0)} / 總計 {s.get('總計應修學分',0)} 學分")
        note = s.get('備註 (模組要求)', '')
        if pd.notna(note) and str(note).strip(): 
            st.markdown(f'<div class="note-box"><b>📝 備註：</b><br>{note}</div>', unsafe_allow_html=True)
    
    for cat in ["必修", "選修"]:
        cat_df = prog_data[prog_data["科目類別"] == cat]
        if not cat_df.empty:
            st.subheader(f"📍 {cat}課程")
            for mod in cat_df["模組名稱"].unique():
                mod_df = cat_df[cat_df["模組名稱"] == mod]
                st.markdown(f'<div class="module-card"><div class="module-title">🔹 {mod}</div>', unsafe_allow_html=True)
                st.dataframe(mod_df[['課程代碼', '課程名稱', '學分數', '課程認抵範圍', '備註']], use_container_width=True, hide_index=True)
                st.markdown('</div>', unsafe_allow_html=True)

# --- TAB 2: 依課程反查學程 ---
with tab_search:
    st.header("🔍 依課程名稱反查所屬學程")
    query = st.text_input("請輸入課程關鍵字 (例如：微積分)")
    if query:
        search_res = df_courses[df_courses["課程名稱"].str.contains(query, case=False, na=False)]
        if not search_res.empty:
            st.dataframe(search_res[['學院', '學程名稱', '適用年度', '科目類別', '課程代碼', '課程名稱']].drop_duplicates(), use_container_width=True, hide_index=True)
        else:
            st.warning("查無相關課程。")

# --- TAB 3: 學程完成度自動比對 (結構化計算修正) ---
with tab_audit:
    st.header("🎓 學程達成度排行 (模組結構化計算)")
    uploaded_file = st.file_uploader("上傳您的成績單 Excel", type=["xlsx"])

    if uploaded_file:
        try:
            user_df = pd.read_excel(uploaded_file, dtype=str).dropna(how='all')
            user_df.columns = user_df.columns.str.strip()
            user_df['學分數'] = pd.to_numeric(user_df['學分數'], errors='coerce').fillna(0.0)
            user_df['及格'] = user_df['成績'].apply(check_passing)
            passed_df = user_df[user_df['及格']]

            audit_results = []
            for _, p_sum_row in df_summary.iterrows():
                p_name = p_sum_row['學程名稱']
                p_year = p_sum_row['適用年度']
                p_courses = df_courses[(df_courses['學程名稱'] == p_name) & (df_courses['適用年度'] == p_year)].copy()
                if p_courses.empty: continue
                
                # 1. 完成判定與互斥處理
                p_courses['已完成'] = p_courses.apply(lambda r: check_course_completion(r, passed_df), axis=1)
                p_courses['採計學分'] = p_courses['學分數']
                p_courses['互斥標記'] = ""
                used_mutex = set()
                for idx, row in p_courses.iterrows():
                    mutex = str(row.get('互斥代碼', '')).strip()
                    if row['已完成'] and mutex and mutex not in ['nan', '']:
                        if mutex in used_mutex:
                            p_courses.at[idx, '採計學分'] = 0.0
                            p_courses.at[idx, '互斥標記'] = f"與 {mutex} 互斥"
                        else:
                            used_mutex.add(mutex)

                # 2. 結構化百分比計算：(模組 A 達成% + 模組 B 達成% ...) / 模組總數
                mod_list = p_courses['模組名稱'].unique()
                total_mod_pct = 0
                for mod in mod_list:
                    mod_data = p_courses[p_courses['模組名稱'] == mod]
                    mod_done = mod_data.loc[mod_data['已完成'], '採計學分'].sum()
                    mod_req = parse_required_credits(mod)
                    
                    if mod_req > 0:
                        total_mod_pct += min(mod_done / mod_req, 1.0)
                    else:
                        total_mod_pct += 1.0 if mod_done > 0 else 0.0
                
                final_pct = total_mod_pct / len(mod_list) if len(mod_list) > 0 else 0.0
                total_actual_done = p_courses.loc[p_courses['已完成'], '採計學分'].sum()
                
                audit_results.append({
                    "name": p_name, "year": p_year, "pct": final_pct, 
                    "done": total_actual_done, "goal": float(p_sum_row['總計應修學分']),
                    "note": p_sum_row.get('備註 (模組要求)', ''),
                    "details": p_courses
                })

            audit_results.sort(key=lambda x: x['pct'], reverse=True)

            for res in audit_results:
                with st.container():
                    st.markdown('<div class="prog-container">', unsafe_allow_html=True)
                    col_t, col_p = st.columns([4, 1])
                    col_t.markdown(f"#### 🎓 {res['name']} ({res['year']}年度)")
                    col_p.markdown(f"### {int(res['pct']*100)}%")
                    st.progress(res['pct'])
                    
                    with st.expander(f"🔍 查看明細 (結構化完成度：{int(res['pct']*100)}%)"):
                        for mod_name in res['details']['模組名稱'].unique():
                            mod_data = res['details'][res['details']['模組名稱'] == mod_name]
                            mod_done = mod_data.loc[mod_data['已完成'], '採計學分'].sum()
                            req_credits = parse_required_credits(mod_name)
                            is_satisfied = (mod_done >= req_credits) if req_credits > 0 else (mod_done > 0)
                            
                            st.markdown(f"""
                                <div class="module-summary-box {'status-done' if is_satisfied else 'status-pending'}">
                                    <span>{'✅' if is_satisfied else '⌛'} {mod_name}</span>
                                    <span>{int(mod_done)} / {int(req_credits)} 學分</span>
                                </div>
                            """, unsafe_allow_html=True)
                            
                            disp_df = mod_data.copy()
                            disp_df['狀態'] = disp_df.apply(lambda r: "✅ 已達成" if r['已完成'] and r['採計學分']>0 else ("⚠️ 互斥不採計" if r['互斥標記'] else "❌ 未達成"), axis=1)
                            st.table(disp_df[['科目類別', '課程代碼', '課程名稱', '學分數', '課程認抵範圍', '備註', '互斥代碼', '狀態']])
                    st.markdown('</div>', unsafe_allow_html=True)
        except Exception as e:
            st.error(f"分析失敗：{e}")