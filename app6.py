import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="學分學程查詢系統", layout="wide")

# --- 1. 資料讀取與樣式設定 ---
@st.cache_data
def load_all_data(file_path):
    try:
        all_sheets = pd.read_excel(file_path, sheet_name=None, dtype=str)
        df_courses = all_sheets["科目表"]
        df_courses['學分數'] = pd.to_numeric(df_courses['學分數'], errors='coerce').fillna(0)
        
        # 格式化規則提醒
        def format_rules(row):
            rules = []
            if pd.notna(row.get('選修群組')) and str(row['選修群組']).strip() != "":
                rules.append(f"📦 {row['選修群組']} ({row.get('群組要求', '無要求')})")
            if pd.notna(row.get('互斥代碼')) and str(row['互斥代碼']).strip() != "":
                rules.append(f"⚠️ 互斥代碼: {row['互斥代碼']}")
            return " / ".join(rules) if rules else "-"
        
        df_courses['規則提醒'] = df_courses.apply(format_rules, axis=1)
        
        # 處理模組欄位：若空白則填入"一般科目"
        if '模組名稱' not in df_courses.columns:
            df_courses['模組名稱'] = "一般科目"
        else:
            df_courses['模組名稱'] = df_courses['模組名稱'].fillna("一般科目").replace("", "一般科目")
            
        df_summary = all_sheets.get("學程規範總額", pd.DataFrame())
        return df_courses, df_summary
    except Exception as e:
        st.error(f"讀取失敗：{e}")
        st.stop()

# 樣式函數
def highlight_keywords(val):
    if st.session_state.get('last_query') and st.session_state.last_query.lower() in str(val).lower():
        return 'background-color: #FFF9C4; color: #E65100; font-weight: bold'
    return ''

def highlight_scope(val):
    target_keywords = ["通識", "系外", "跨系", "自由選修"]
    for key in target_keywords:
        if key in str(val):
            return 'background-color: #E8F5E9; color: #2E7D32; font-weight: bold'
    return ''

# --- 2. 核心功能：顯示分組表格 ---
def display_grouped_tables(df, is_search_mode=False):
    """
    將傳入的 DataFrame 依照 必選修 -> 模組名稱 進行分群顯示
    """
    categories = ["必修", "選修"]
    for cat in categories:
        cat_df = df[df["科目類別"] == cat]
        if not cat_df.empty:
            st.markdown(f"### 📍 {cat}科目")
            
            # 依照模組分群
            modules = cat_df["模組名稱"].unique()
            for mod in modules:
                mod_df = cat_df[cat_df["模組名稱"] == mod]
                with st.expander(f"🔹 模組：{mod} (共 {len(mod_df)} 門課)", expanded=True):
                    # 套用樣式
                    cols_to_show = ['課程代碼', '課程名稱', '學分數', '課程認抵範圍', '規則提醒']
                    styled_df = mod_df[cols_to_show].style \
                        .applymap(highlight_keywords, subset=['課程代碼', '課程名稱']) \
                        .applymap(highlight_scope, subset=['課程認抵範圍'])
                    
                    st.dataframe(styled_df, use_container_width=True, hide_index=True)
        else:
            if not is_search_mode:
                st.write(f"（無{cat}科目資料）")

# --- 3. 介面初始化 ---
DATA_FILE = "master_data.xlsx"
if os.path.exists(DATA_FILE):
    df_courses, df_summary = load_all_data(DATA_FILE)
else:
    st.error("找不到 master_data.xlsx")
    st.stop()

tab_browse, tab_search = st.tabs(["📂 按學程瀏覽", "🔍 依課程反查學程"])

# --- TAB 1: 按學程瀏覽 ---
with tab_browse:
    st.header("學分學程規範查詢")
    
    def reset_browse():
        st.session_state.b1 = "全部"

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
        st.button("🔄 重設篩選", on_click=reset_browse, use_container_width=True, key="btn_res_b")

    summary = df_summary[(df_summary["學程名稱"] == sel_prog) & (df_summary["適用年度"] == sel_year)]
    if not summary.empty:
        s = summary.iloc[0]
        st.success(f"🎓 **畢業門檻：** 必修 {s.get('必修總學分',0)} / 選修 {s.get('選修總學分',0)} / 總計 {s.get('總計應修學分',0)} 學分")
        if pd.notna(s.get('備註 (模組要求)')): st.info(f"💡 **備註：** {s['備註 (模組要求)']}")
    
    current_prog_df = df_courses[(df_courses["學程名稱"] == sel_prog) & (df_courses["適用年度"] == sel_year)]
    display_grouped_tables(current_prog_df)


# --- TAB 2: 依課程反查學程 ---
with tab_search:
    st.header("🔍 搜尋課程找學程")
    
    def clear_search():
        st.session_state.s1 = ""
        st.session_state.last_query = ""
        if 's2' in st.session_state: st.session_state.s2 = "--- 請選擇 ---"

    col_search, col_clear = st.columns([4, 1])
    with col_search:
        course_query = st.text_input("📝 請輸入課程名稱或代碼", key="s1")
        st.session_state.last_query = course_query

    with col_clear:
        st.markdown("<br>", unsafe_allow_html=True)
        st.button("🧹 清除條件", on_click=clear_search, use_container_width=True, key="btn_clear_s")

    if course_query:
        search_results = df_courses[
            df_courses["課程名稱"].str.contains(course_query, case=False, na=False) |
            df_courses["課程代碼"].str.contains(course_query, case=False, na=False)
        ].copy()
        
        if not search_results.empty:
            st.write(f"以下學程包含關鍵字「{course_query}」：")
            mini_df = search_results[['學院', '學程名稱', '適用年度', '課程代碼', '課程名稱']].drop_duplicates()
            styled_mini = mini_df.style.applymap(highlight_keywords, subset=['課程代碼', '課程名稱'])
            st.dataframe(styled_mini, use_container_width=True, hide_index=True)
            
            st.divider()
            
            st.subheader("📖 查看該學程完整資訊")
            options = mini_df.apply(lambda x: f"{x['學程名稱']} ({x['適用年度']})", axis=1).unique().tolist()
            selected_option = st.selectbox("選擇學程展開詳情：", ["--- 請選擇 ---"] + options, key="s2")
            
            if selected_option != "--- 請選擇 ---":
                p_name = selected_option.split(" (")[0]
                p_year = selected_option.split(" (")[1].replace(")", "")
                detail_df = df_courses[(df_courses["學程名稱"] == p_name) & (df_courses["適用年度"] == p_year)]
                detail_summary = df_summary[(df_summary["學程名稱"] == p_name) & (df_summary["適用年度"] == p_year)]
                
                with st.expander(f"📌 {selected_option} 完整清單", expanded=True):
                    if not detail_summary.empty:
                        ds = detail_summary.iloc[0]
                        st.info(f"**門檻：** 必修 {ds.get('必修總學分',0)} / 選修 {ds.get('選修總學分',0)} / 總計 {ds.get('總計應修學分',0)} 學分")
                    
                    # 使用相同的分組顯示函數
                    display_grouped_tables(detail_df, is_search_mode=True)
        else:
            st.warning("查無相關課程。")