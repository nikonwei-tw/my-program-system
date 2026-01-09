import streamlit as st
import pandas as pd
import os
import io

st.set_page_config(page_title="學分學程查詢系統", layout="wide")

# --- 1. 自定義 CSS (保留原本，並新增比對用的卡片樣式) ---
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
    /* 新增比對分頁專用的樣式 */
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

# --- 2. 資料讀取與處理 (維持原邏輯，僅加強學分轉數字功能以供計算) ---
@st.cache_data
def load_all_data(file_path):
    try:
        all_sheets = pd.read_excel(file_path, sheet_name=None, dtype=str)
        df_courses = all_sheets["科目表"]
        df_courses['學分數'] = pd.to_numeric(df_courses['學分數'], errors='coerce').fillna(0)
        
        # 規則提醒邏輯 (保留您原本的設計)
        def format_rules(row):
            rules = []
            if pd.notna(row.get('互斥代碼')) and str(row['互斥代碼']).strip() != "":
                rules.append(f"⚠️ 互斥代碼: {row['互斥代碼']}")
            return " / ".join(rules) if rules else "-"
        
        df_courses['規則提醒'] = df_courses.apply(format_rules, axis=1)
        df_courses['模組名稱'] = df_courses.get('模組名稱', "一般科目").fillna("一般科目").replace("", "一般科目")
        
        df_summary = all_sheets.get("學程規範總額", pd.DataFrame())
        # 為了比對計算，確保總額表學分為數字
        for col in ['必修總學分', '選修總學分', '總計應修學分']:
            if col in df_summary.columns:
                df_summary[col] = pd.to_numeric(df_summary[col], errors='coerce').fillna(0.0)
        
        return df_courses, df_summary
    except Exception as e:
        st.error(f"讀取 Excel 失敗：{e}")
        st.stop()

# --- 原本的高亮與下載功能 (完全不動) ---
def highlight_keywords(val):
    query = st.session_state.get('last_query', "")
    if query and query.lower() in str(val).lower():
        return 'background-color: #FFF9C4; color: #E65100; font-weight: bold'
    return ''

def highlight_scope(val):
    for key in ["通識", "系外", "跨系", "自由選修"]:
        if key in str(val):
            return 'background-color: #E8F5E9; color: #2E7D32; font-weight: bold'
    return ''

def convert_df_to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='學程課表')
    return output.getvalue()

# --- 比對功能所需的輔助函數 ---
def check_passing(grade):
    grade_str = str(grade).strip().upper()
    if grade_str in ["通過", "及格", "P", "PASS"]: return True
    try:
        return float(grade_str) >= 60
    except:
        return False

def get_template():
    df_temp = pd.DataFrame(columns=['課程代碼', '課程名稱', '學分數', '開課單位', '成績'])
    df_temp.loc[0] = ['範例代碼', '範例課名', '3', '資工系', '85']
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_temp.to_excel(writer, index=False)
    return output.getvalue()

# --- 主程式 ---
DATA_FILE = "master_data.xlsx"
if os.path.exists(DATA_FILE):
    df_courses, df_summary = load_all_data(DATA_FILE)
else:
    st.error(f"找不到檔案: {DATA_FILE}")
    st.stop()

# 新增第三個分頁標籤
tab_browse, tab_search, tab_audit = st.tabs(["📂 按學程瀏覽", "🔍 依課程反查學程", "🎓 畢業達成度比對"])

# --- TAB 1: 保持您的原本程式碼 ---
with tab_browse:
    st.header("學分學程規範查詢")
    col_b1, col_b2, col_b3, col_reset = st.columns([2, 2, 2, 1])
    with col_b1:
        college_list = ["全部"] + sorted(df_courses["學院"].dropna().unique().tolist())
        sel_col = st.selectbox("1. 選擇學院", college_list, key="b1")
    with col_b2:
        temp_p = df_courses if sel_col == "全部" else df_courses[df_courses["學院"] == sel_col]
        prog_list = sorted(temp_p["學程名稱"].dropna().unique().tolist())
        sel_prog = st.selectbox("2. 選擇學程", prog_list, key="b2")
    with col_b3:
        year_list = sorted(df_courses[df_courses["學程名稱"]==sel_prog]["適用年度"].unique(), reverse=True)
        sel_year = st.selectbox("3. 選擇年度", year_list, key="b3")
    with col_reset:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("🔄 重設篩選", use_container_width=True):
            st.session_state.b1 = "全部"
            st.rerun()

    prog_data = df_courses[(df_courses["學程名稱"] == sel_prog) & (df_courses["適用年度"] == sel_year)]
    st.divider()
    head_col, down_col = st.columns([4, 1])
    with head_col:
        st.subheader(f"🎓 {sel_prog} ({sel_year}年度)")
    with down_col:
        excel_file = convert_df_to_excel(prog_data)
        st.download_button(label="📥 下載完整 Excel", data=excel_file, file_name=f"{sel_prog}_{sel_year}.xlsx", mime="application/vnd.ms-excel", use_container_width=True)

    summary = df_summary[(df_summary["學程名稱"] == sel_prog) & (df_summary["適用年度"] == sel_year)]
    if not summary.empty:
        s = summary.iloc[0]
        st.success(f"**畢業門檻：** 必修 {s.get('必修總學分',0)} / 選修 {s.get('選修總學分',0)} / 總計 {s.get('總計應修學分',0)} 學分")
        if pd.notna(s.get('備註 (模組要求)')):
            st.info(f"💡 **備註：** {s['備註 (模組要求)']}")

    # 顯示表格 (維持您的原樣)
    for cat in ["必修", "選修"]:
        cat_df = prog_data[prog_data["科目類別"] == cat]
        if not cat_df.empty:
            st.markdown(f"### 📍 {cat}課程清單")
            for mod in cat_df["模組名稱"].unique():
                mod_df = cat_df[cat_df["模組名稱"] == mod]
                st.markdown(f'<div class="module-card"><div class="module-title">🔹 {mod}</div><div class="module-info">ℹ️ 共有 {len(mod_df)} 門課，總開放 {mod_df["學分數"].sum()} 學分</div>', unsafe_allow_html=True)
                cols = ['課程代碼', '課程名稱', '學分數', '課程認抵範圍', '規則提醒']
                st.dataframe(mod_df[cols].style.applymap(highlight_keywords, subset=['課程代碼', '課程名稱']).applymap(highlight_scope, subset=['課程認抵範圍']), use_container_width=True, hide_index=True)
                st.markdown('</div>', unsafe_allow_html=True)

# --- TAB 2: 保持您的原本程式碼 ---
with tab_search:
    st.header("🔍 搜尋課程找學程")
    col_search, col_clear = st.columns([4, 1])
    with col_search:
        course_query = st.text_input("📝 請輸入課程名稱或代碼關鍵字", key="s1")
        st.session_state.last_query = course_query
    with col_clear:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("🧹 清除條件", use_container_width=True):
            st.session_state.s1 = ""
            st.session_state.last_query = ""
            st.rerun()

    if course_query:
        search_res = df_courses[df_courses["課程名稱"].str.contains(course_query, case=False, na=False) | df_courses["課程代碼"].str.contains(course_query, case=False, na=False)].copy()
        if not search_res.empty:
            mini_df = search_res[['學院', '學程名稱', '適用年度', '課程代碼', '課程名稱']].drop_duplicates()
            st.dataframe(mini_df.style.applymap(highlight_keywords, subset=['課程代碼', '課程名稱']), use_container_width=True, hide_index=True)
            st.divider()
            st.subheader("📖 查看完整學程資訊")
            prog_options = mini_df.apply(lambda x: f"{x['學程名稱']} ({x['適用年度']})", axis=1).unique().tolist()
            selected_prog = st.selectbox("請選擇上述搜尋結果中的學程：", ["--- 請選擇 ---"] + prog_options, key="s2")
            if selected_prog != "--- 請選擇 ---":
                p_n = selected_prog.split(" (")[0]
                p_y = selected_prog.split(" (")[1].replace(")", "")
                d_data = df_courses[(df_courses["學程名稱"] == p_n) & (df_courses["適用年度"] == p_y)]
                # 重複 TAB 1 的顯示邏輯 (維持原狀)
                for cat in ["必修", "選修"]:
                    cat_df = d_data[d_data["科目類別"] == cat]
                    if not cat_df.empty:
                        st.markdown(f"### 📍 {cat}課程清單")
                        for mod in cat_df["模組名稱"].unique():
                            mod_df = cat_df[cat_df["模組名稱"] == mod]
                            st.markdown(f'<div class="module-card"><div class="module-title">🔹 {mod}</div>', unsafe_allow_html=True)
                            st.dataframe(mod_df[['課程代碼', '課程名稱', '學分數', '課程認抵範圍']].style.applymap(highlight_keywords, subset=['課程代碼', '課程名稱']), use_container_width=True, hide_index=True)
                            st.markdown('</div>', unsafe_allow_html=True)
        else:
            st.warning("查無相關課程。")

# --- TAB 3: 新增的比對功能 (完全獨立，不影響前兩頁) ---
with tab_audit:
    st.header("🎓 學程完成度自動比對")
    st.info("請下載範本填寫成績後上傳，系統將自動比對所有學程的達成進度。")
    
    col1, col2 = st.columns([1, 2])
    with col1:
        st.download_button("📥 下載成績單範本", data=get_template(), file_name="成績單範本.xlsx")
    with col2:
        uploaded_file = st.file_uploader("上傳您的成績單 (Excel)", type=["xlsx"])

    available_progs = sorted(df_summary['學程名稱'].dropna().unique().tolist())
    selected_progs = st.multiselect("勾選欲比對的學程 (預設為全選)", available_progs, default=available_progs)

    if uploaded_file:
        try:
            user_df = pd.read_excel(uploaded_file, dtype=str).dropna(how='all')
            user_df['學分數'] = pd.to_numeric(user_df['學分數'], errors='coerce').fillna(0.0)
            user_df['及格'] = user_df['成績'].apply(check_passing)
            passed_df = user_df[user_df['及格']]
            
            st.divider()
            for p_name in selected_progs:
                # 抓取該學程最新規範
                p_sum_all = df_summary[df_summary['學程名稱'] == p_name].sort_values('適用年度', ascending=False)
                if p_sum_all.empty: continue
                p_sum = p_sum_all.iloc[0]
                p_courses = df_courses[(df_courses['學程名稱'] == p_name) & (df_courses['適用年度'] == p_sum['適用年度'])].copy()
                
                # 比對函數
                def check_done(row):
                    match = passed_df[passed_df['課程代碼'].str.strip() == str(row['課程代碼']).strip()]
                    return match['學分數'].sum() >= float(row['學分數']) if not match.empty else False

                p_courses['已完成'] = p_courses.apply(check_done, axis=1)
                d_req = float(p_courses[p_courses['已完成'] & (p_courses['科目類別'] == '必修')]['學分數'].sum())
                d_elec = float(p_courses[p_courses['已完成'] & (p_courses['科目類別'] == '選修')]['學分數'].sum())
                g_req, g_elec, g_total = float(p_sum['必修總學分']), float(p_sum['選修總學分']), float(p_sum['總計應修學分'])
                
                pct = min((d_req + d_elec) / g_total, 1.0) if g_total > 0 else 0.0

                # 顯示卡片
                st.markdown('<div class="prog-container">', unsafe_allow_html=True)
                ca, cb = st.columns([4, 1])
                ca.markdown(f"### 🎓 {p_name} ({p_sum['適用年度']})")
                ca.progress(pct)
                cb.markdown(f"## {int(pct*100)}%")
                
                m1, m2, m3 = st.columns(3)
                m1.metric("必修進度", f"{int(d_req)}/{int(g_req)}")
                m2.metric("選修進度", f"{int(d_elec)}/{int(g_elec)}")
                
                s_req, s_elec = max(0, g_req - d_req), max(0, g_elec - d_elec)
                if (d_req + d_elec) >= g_total and s_req == 0:
                    st.success("🎉 已達成門檻！")
                else:
                    st.warning(f"⚠️ 尚缺：必修 {int(s_req)} / 選修 {int(s_elec)} 學分")
                st.markdown('</div>', unsafe_allow_html=True)
        except Exception as e:
            st.error(f"分析出錯：{e}")