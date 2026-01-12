import streamlit as st
import pandas as pd
import os
import io

# --- 0. 頁面設定 ---
st.set_page_config(page_title="學分學程查詢與比對系統", layout="wide")

# --- 1. 自定義 CSS ---
st.markdown("""
    <style>
    .module-card {
        background-color: #f8f9fa;
        border-radius: 10px;
        padding: 18px;
        margin-bottom: 20px;
        border-left: 6px solid #007bff;
        box-shadow: 2px 2px 8px rgba(0,0,0,0.08);
    }
    .module-title { color: #0056b3; font-weight: bold; font-size: 1.15rem; margin-bottom: 5px; }
    .prog-container {
        background-color: #ffffff;
        border: 1px solid #e0e0e0;
        border-radius: 10px;
        padding: 20px;
        margin-bottom: 25px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
    }
    </style>
    """, unsafe_allow_html=True)

# --- 2. 資料讀取與強化處理 ---
@st.cache_data
def load_all_data(file_path):
    try:
        all_sheets = pd.read_excel(file_path, sheet_name=None, dtype=str)
        
        # A. 科目表處理
        df_courses = all_sheets["科目表"]
        df_courses.columns = df_courses.columns.str.strip() # 清理欄位空格
        
        df_courses['學分數'] = pd.to_numeric(df_courses['學分數'], errors='coerce').fillna(0)
        df_courses['模組名稱'] = df_courses.get('模組名稱', "一般科目").fillna("一般科目")
        
        # 修正 K860 問題：建立代碼對名稱的對照表
        prog_map = df_courses[['學程代碼', '學程名稱']].drop_duplicates().set_index('學程代碼')['學程名稱'].to_dict()
        
        # B. 規範總額表處理
        df_summary = all_sheets.get("學程規範總額", pd.DataFrame())
        df_summary.columns = df_summary.columns.str.strip()
        
        if not df_summary.empty and '學程代碼' in df_summary.columns:
            # 強制將總額表的名稱修正為科目表中的正確名稱
            df_summary['學程名稱'] = df_summary['學程代碼'].map(prog_map).fillna(df_summary.get('學程名稱', '未知學程'))
            
            for col in ['必修總學分', '選修總學分', '總計應修學分']:
                if col in df_summary.columns:
                    df_summary[col] = pd.to_numeric(df_summary[col], errors='coerce').fillna(0.0)
        
        return df_courses, df_summary, prog_map
    except Exception as e:
        st.error(f"讀取 Excel 失敗。錯誤：{e}"); st.stop()

# --- 3. 狀態重設函數 (修復錯誤的關鍵) ---
def reset_browser_filters():
    # 透過回呼函數修改 session_state 是安全且被允許的
    st.session_state.b1 = "全部"

# --- 4. 輔助功能 ---
def check_passing(grade):
    g = str(grade).strip().upper()
    if g in ["通過", "及格", "P", "PASS"]: return True
    try: return float(g) >= 60
    except: return False

def check_course_completion(req_row, passed_df):
    """判斷該科目是否認抵成功 (支援多單位 D51, D89)"""
    matches = passed_df[passed_df['課程代碼'].str.strip() == str(req_row['課程代碼']).strip()]
    if matches.empty: return False

    allowed_val = str(req_row.get('認抵單位代碼', 'ANY')).strip().upper()
    if allowed_val in ["ANY", "NAN", "", "全部"]:
        return matches['學分數'].sum() >= float(req_row['學分數'])
    
    allowed_list = [d.strip() for d in allowed_val.split(',')]
    if '開課單位代碼' in matches.columns:
        valid_matches = matches[matches['開課單位代碼'].str.upper().isin(allowed_list)]
        return valid_matches['學分數'].sum() >= float(req_row['學分數'])
    return False

# --- 5. 主程式執行 ---
DATA_FILE = "master_data.xlsx"
if os.path.exists(DATA_FILE):
    df_courses, df_summary, prog_map = load_all_data(DATA_FILE)
else:
    st.error("找不到 master_data.xlsx"); st.stop()

tab_browse, tab_search, tab_audit = st.tabs(["📂 按學程瀏覽", "🔍 依課程反查學程", "🎓 畢業達成度比對"])

# --- TAB 1: 瀏覽功能 ---
with tab_browse:
    st.header("學分學程規範查詢")
    
    # 初始 session_state
    if "b1" not in st.session_state:
        st.session_state.b1 = "全部"

    col_b1, col_b2, col_b3, col_reset = st.columns([2, 2, 2, 1])
    with col_b1:
        college_list = ["全部"] + sorted(df_courses["學院"].dropna().unique().tolist())
        sel_col = st.selectbox("1. 選擇學院", college_list, key="b1")
    
    with col_b2:
        temp_p = df_courses if sel_col == "全部" else df_courses[df_courses["學院"] == sel_col]
        prog_list = sorted(temp_p["學程名稱"].dropna().unique().tolist())
        sel_prog = st.selectbox("2. 選擇學程", prog_list)
    
    with col_b3:
        year_list = sorted(df_courses[df_courses["學程名稱"]==sel_prog]["適用年度"].unique(), reverse=True)
        sel_year = st.selectbox("3. 選擇年度", year_list)
    
    with col_reset:
        st.markdown("<br>", unsafe_allow_html=True)
        # 關鍵：使用 on_click 觸發重設函數
        st.button("🔄 重設篩選", on_click=reset_browser_filters, use_container_width=True)

    # 顯示內容邏輯...
    prog_data = df_courses[(df_courses["學程名稱"] == sel_prog) & (df_courses["適用年度"] == sel_year)]
    summary = df_summary[(df_summary["學程名稱"] == sel_prog) & (df_summary["適用年度"] == sel_year)]
    
    if not summary.empty:
        s = summary.iloc[0]
        st.success(f"**畢業門檻：** 必修 {s.get('必修總學分',0)} / 選修 {s.get('選修總學分',0)} / 總計 {s.get('總計應修學分',0)} 學分")

    for cat in ["必修", "選修"]:
        cat_df = prog_data[prog_data["科目類別"] == cat]
        if not cat_df.empty:
            st.markdown(f"### 📍 {cat}課程清單")
            for mod in cat_df["模組名稱"].unique():
                mod_df = cat_df[cat_df["模組名稱"] == mod]
                st.markdown(f'<div class="module-card"><div class="module-title">🔹 {mod}</div>', unsafe_allow_html=True)
                # 容錯過濾欄位
                target_cols = ['課程代碼', '課程名稱', '學分數', '課程認抵範圍', '認抵單位代碼']
                available_cols = [c for c in target_cols if c in mod_df.columns]
                st.dataframe(mod_df[available_cols], use_container_width=True, hide_index=True)
                st.markdown('</div>', unsafe_allow_html=True)

# --- TAB 2: 搜尋功能 (簡化版) ---
with tab_search:
    st.header("🔍 搜尋課程找學程")
    course_query = st.text_input("📝 請輸入課程名稱或代碼關鍵字", key="search_input")
    if course_query:
        search_res = df_courses[df_courses["課程名稱"].str.contains(course_query, case=False, na=False) | 
                                df_courses["課程代碼"].str.contains(course_query, case=False, na=False)].copy()
        if not search_res.empty:
            mini_df = search_res[['學院', '學程名稱', '適用年度', '課程代碼', '課程名稱']].drop_duplicates()
            st.dataframe(mini_df, use_container_width=True, hide_index=True)
        else:
            st.warning("查無相關課程。")

# --- TAB 3: 比對功能 (完整版) ---
with tab_audit:
    st.header("🎓 學程完成度自動比對")
    uploaded_file = st.file_uploader("上傳您的成績單 (Excel)", type=["xlsx"])
    
    all_audit_progs = sorted(df_summary['學程名稱'].dropna().unique().tolist())
    selected_audit_progs = st.multiselect("勾選欲比對的學程", all_audit_progs, default=all_audit_progs)

    if uploaded_file:
        try:
            user_df = pd.read_excel(uploaded_file, dtype=str).dropna(how='all')
            user_df.columns = user_df.columns.str.strip()
            user_df['學分數'] = pd.to_numeric(user_df['學分數'], errors='coerce').fillna(0.0)
            user_df['及格'] = user_df['成績'].apply(check_passing)
            passed_df = user_df[user_df['及格']]

            for p_name in selected_audit_progs:
                p_sum_row = df_summary[df_summary['學程名稱'] == p_name].sort_values('適用年度', ascending=False).iloc[0]
                p_courses = df_courses[(df_courses['學程名稱'] == p_name) & (df_courses['適用年度'] == p_sum_row['適用年度'])].copy()
                
                p_courses['已完成'] = p_courses.apply(lambda r: check_course_completion(r, passed_df), axis=1)
                
                # 學分統計
                done_total = p_courses[p_courses['已完成']]['學分數'].sum()
                goal_total = float(p_sum_row['總計應修學分'])
                pct = min(done_total / goal_total, 1.0) if goal_total > 0 else 0.0

                st.markdown('<div class="prog-container">', unsafe_allow_html=True)
                ca, cb = st.columns([4, 1])
                ca.markdown(f"### 🎓 {p_name} ({p_sum_row['適用年度']}年度)")
                ca.progress(pct)
                cb.markdown(f"## {int(pct*100)}%")
                st.write(f"已達成學分：{int(done_total)} / 門檻：{int(goal_total)}")
                with st.expander("查看明細"):
                    st.table(p_courses[['科目類別', '課程代碼', '課程名稱', '學分數', '已完成']].replace({True: "✅", False: "❌"}))
                st.markdown('</div>', unsafe_allow_html=True)
        except Exception as e:
            st.error(f"分析出錯：{e}")