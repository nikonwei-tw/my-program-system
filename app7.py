import streamlit as st
import pandas as pd
import os
import io

# --- 0. 頁面設定 ---
st.set_page_config(page_title="學分學程查詢與比對系統", layout="wide")

# --- 1. 自定義 CSS (完全保留並擴充) ---
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
    .module-title {
        color: #0056b3;
        font-weight: bold;
        font-size: 1.15rem;
        margin-bottom: 5px;
    }
    .module-info {
        font-size: 0.85rem;
        color: #666;
        margin-bottom: 12px;
    }
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

# --- 2. 資料讀取與核心處理邏輯 ---
@st.cache_data
def load_all_data(file_path):
    try:
        all_sheets = pd.read_excel(file_path, sheet_name=None, dtype=str)
        
        # A. 科目表處理
        df_courses = all_sheets["科目表"]
        df_courses['學分數'] = pd.to_numeric(df_courses['學分數'], errors='coerce').fillna(0)
        df_courses['模組名稱'] = df_courses.get('模組名稱', "一般科目").fillna("一般科目")
        
        # 修正 K860 問題：建立代碼對名稱的對照表
        prog_map = df_courses[['學程代碼', '學程名稱']].drop_duplicates().set_index('學程代碼')['學程名稱'].to_dict()
        
        # B. 規範總額表處理
        df_summary = all_sheets.get("學程規範總額", pd.DataFrame())
        if not df_summary.empty and '學程代碼' in df_summary.columns:
            # 強制將總額表的名稱修正為科目表中的正確名稱 (解決 K860 顯示問題)
            df_summary['學程名稱'] = df_summary['學程代碼'].map(prog_map).fillna(df_summary['學程名稱'])
            
            for col in ['必修總學分', '選修總學分', '總計應修學分']:
                if col in df_summary.columns:
                    df_summary[col] = pd.to_numeric(df_summary[col], errors='coerce').fillna(0.0)
        
        return df_courses, df_summary, prog_map
    except Exception as e:
        st.error(f"讀取 Excel 失敗，請檢查檔案格式與工作表名稱。錯誤：{e}")
        st.stop()

# --- 3. 輔助函數 ---
def highlight_keywords(val):
    query = st.session_state.get('last_query', "")
    return 'background-color: #FFF9C4; color: #E65100; font-weight: bold' if query and query.lower() in str(val).lower() else ''

def highlight_scope(val):
    for key in ["通識", "系外", "跨系", "自由選修"]:
        if key in str(val): return 'background-color: #E8F5E9; color: #2E7D32; font-weight: bold'
    return ''

def check_passing(grade):
    g = str(grade).strip().upper()
    if g in ["通過", "及格", "P", "PASS"]: return True
    try: return float(g) >= 60
    except: return False

def get_template():
    df_temp = pd.DataFrame(columns=['課程代碼', '課程名稱', '學分數', '開課單位代碼', '成績'])
    df_temp.loc[0] = ['CS101', '範例程式設計', '3', 'D51', '85']
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_temp.to_excel(writer, index=False)
    return output.getvalue()

def check_course_completion(req_row, passed_df):
    """核心比對邏輯：檢查單一科目是否完成認抵"""
    # 1. 找代碼相同的課
    matches = passed_df[passed_df['課程代碼'].str.strip() == str(req_row['課程代碼']).strip()]
    if matches.empty: return False

    # 2. 檢查認抵單位限制 (支援多單位如 D51, D89)
    # 假設 Excel 欄位名稱為 '認抵單位代碼'，若無則預設為 ANY
    allowed_val = str(req_row.get('認抵單位代碼', 'ANY')).strip().upper()
    if allowed_val in ["ANY", "NAN", "", "全部"]:
        return matches['學分數'].sum() >= float(req_row['學分數'])
    
    # 拆解允許多單位 (D51, D89 -> ['D51', 'D89'])
    allowed_list = [d.strip() for d in allowed_val.split(',')]
    
    # 檢查學生修的單位是否在名單內
    if '開課單位代碼' in matches.columns:
        valid_matches = matches[matches['開課單位代碼'].str.upper().isin(allowed_list)]
        return valid_matches['學分數'].sum() >= float(req_row['學分數'])
    
    return False

# --- 4. 主程式流程 ---
DATA_FILE = "master_data.xlsx"
if os.path.exists(DATA_FILE):
    df_courses, df_summary, prog_map = load_all_data(DATA_FILE)
else:
    st.error(f"找不到資料檔案：{DATA_FILE}，請確認檔案已上傳至伺服器。")
    st.stop()

tab_browse, tab_search, tab_audit = st.tabs(["📂 按學程瀏覽", "🔍 依課程反查學程", "🎓 畢業達成度比對"])

# --- TAB 1: 按學程瀏覽 (保留原始功能) ---
with tab_browse:
    st.header("學分學程規範查詢")
    col_b1, col_b2, col_b3, col_reset = st.columns([2, 2, 2, 1])
    with col_b1:
        college_list = ["全部"] + sorted(df_courses["學院"].dropna().unique().tolist())
        sel_col = st.selectbox("1. 選擇學院", college_list, key="b1")
    with col_b2:
        temp_p = df_courses if sel_col == "全部" else df_courses[df_courses["學院"] == sel_col]
        # 名稱已修正，這裡不會出現 K860
        prog_list = sorted(temp_p["學程名稱"].dropna().unique().tolist())
        sel_prog = st.selectbox("2. 選擇學程", prog_list, key="b2")
    with col_b3:
        year_list = sorted(df_courses[df_courses["學程名稱"]==sel_prog]["適用年度"].unique(), reverse=True)
        sel_year = st.selectbox("3. 選擇年度", year_list, key="b3")
    with col_reset:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("🔄 重設篩選", use_container_width=True):
            st.session_state.b1 = "全部"; st.rerun()

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
                st.dataframe(mod_df[['課程代碼', '課程名稱', '學分數', '課程認抵範圍']], use_container_width=True, hide_index=True)
                st.markdown('</div>', unsafe_allow_html=True)

# --- TAB 2: 依課程反查 (保留原始功能) ---
with tab_search:
    st.header("🔍 搜尋課程找學程")
    course_query = st.text_input("📝 請輸入課程名稱或代碼關鍵字", key="s1")
    st.session_state.last_query = course_query
    
    if course_query:
        search_res = df_courses[df_courses["課程名稱"].str.contains(course_query, case=False, na=False) | 
                                df_courses["課程代碼"].str.contains(course_query, case=False, na=False)].copy()
        if not search_res.empty:
            mini_df = search_res[['學院', '學程名稱', '適用年度', '課程代碼', '課程名稱']].drop_duplicates()
            st.dataframe(mini_df.style.applymap(highlight_keywords, subset=['課程代碼', '課程名稱']), use_container_width=True, hide_index=True)
        else:
            st.warning("查無相關課程。")

# --- TAB 3: 畢業達成度比對 (完整強化版) ---
with tab_audit:
    st.header("🎓 學程完成度自動比對")
    
    # 1. 檔案與設定區
    with st.container():
        c1, c2 = st.columns([1, 2])
        with c1:
            st.download_button("📥 下載成績單範本", data=get_template(), file_name="成績單範本.xlsx", use_container_width=True)
        with c2:
            uploaded_file = st.file_uploader("上傳您的成績單 (Excel)", type=["xlsx"])

    # 2. 單位限制設定 (全域過濾)
    dept_input = st.text_input("📍 全局限定開課單位認抵 (例如輸入 D01,D02。留空則不限單位)", placeholder="例如: D01, D02")
    global_limit_depts = [d.strip().upper() for d in dept_input.split(",")] if dept_input else []

    # 3. 選擇要比對的學程
    all_available_progs = sorted(df_summary['學程名稱'].dropna().unique().tolist())
    selected_audit_progs = st.multiselect("勾選欲比對的學程 (可多選)", all_available_progs, default=all_available_progs)

    if uploaded_file:
        try:
            # 讀取學生資料
            user_df = pd.read_excel(uploaded_file, dtype=str).dropna(how='all')
            user_df['學分數'] = pd.to_numeric(user_df['學分數'], errors='coerce').fillna(0.0)
            user_df['及格'] = user_df['成績'].apply(check_passing)
            
            # 執行全局單位過濾
            if global_limit_depts and '開課單位代碼' in user_df.columns:
                passed_df = user_df[user_df['及格'] & user_df['開課單位代碼'].str.upper().isin(global_limit_depts)]
                st.info(f"💡 已套用單位過濾，符合 {global_limit_depts} 的及格課程共 {len(passed_df)} 門。")
            else:
                passed_df = user_df[user_df['及格']]
            
            st.divider()

            # 4. 開始對每個學程進行比對
            for p_name in selected_audit_progs:
                # 取得該學程最新年度規範
                p_sum_row = df_summary[df_summary['學程名稱'] == p_name].sort_values('適用年度', ascending=False).iloc[0]
                # 取得科目清單
                p_courses = df_courses[(df_courses['學程名稱'] == p_name) & (df_courses['適用年度'] == p_sum_row['適用年度'])].copy()
                
                # 比對每一門課
                p_courses['已完成'] = p_courses.apply(lambda r: check_course_completion(r, passed_df), axis=1)
                
                # 計算學分
                done_req = p_courses[p_courses['已完成'] & (p_courses['科目類別'] == '必修')]['學分數'].sum()
                done_elec = p_courses[p_courses['已完成'] & (p_courses['科目類別'] == '選修')]['學分數'].sum()
                
                goal_req = float(p_sum_row['必修總學分'])
                goal_elec = float(p_sum_row['選修總學分'])
                goal_total = float(p_sum_row['總計應修學分'])
                
                total_done = done_req + done_elec
                progress_pct = min(total_done / goal_total, 1.0) if goal_total > 0 else 0.0

                # 5. 顯示進度卡片
                st.markdown('<div class="prog-container">', unsafe_allow_html=True)
                ca, cb = st.columns([4, 1])
                ca.markdown(f"### 🎓 {p_name} ({p_sum_row['適用年度']}年度)")
                ca.progress(progress_pct)
                cb.markdown(f"## {int(progress_pct*100)}%")
                
                m1, m2, m3 = st.columns(3)
                m1.metric("必修學分", f"{int(done_req)} / {int(goal_req)}")
                m2.metric("選修學分", f"{int(done_elec)} / {int(goal_elec)}")
                m3.metric("總達成度", f"{int(total_done)} / {int(goal_total)}")
                
                with st.expander("🔍 查看詳細比對清單"):
                    # 整理顯示內容
                    display_df = p_courses[['科目類別', '課程代碼', '課程名稱', '學分數', '已完成']].copy()
                    display_df['已完成'] = display_df['已完成'].map({True: "✅ 已達成", False: "❌ 未完成"})
                    st.table(display_df)
                st.markdown('</div>', unsafe_allow_html=True)

        except Exception as e:
            st.error(f"分析過程出錯，請確認上傳的檔案欄位是否正確（課程代碼、學分數、成績）。錯誤：{e}")