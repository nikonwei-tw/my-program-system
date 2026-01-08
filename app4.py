import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="學分學程查詢系統", layout="wide")

# --- 1. 資料讀取與規則格式化函數 ---
@st.cache_data
def load_all_data(file_path):
    try:
        all_sheets = pd.read_excel(file_path, sheet_name=None, dtype=str)
        df_courses = all_sheets["科目表"]
        
        # 轉換學分為數字
        df_courses['學分數'] = pd.to_numeric(df_courses['學分數'], errors='coerce').fillna(0)
        
        # --- 自動產生「規則提醒」欄位 ---
        def format_rules(row):
            rules = []
            if pd.notna(row.get('選修群組')) and str(row['選修群組']).strip() != "":
                rules.append(f"📦 {row['選修群組']} ({row.get('群組要求', '無要求')})")
            if pd.notna(row.get('互斥代碼')) and str(row['互斥代碼']).strip() != "":
                rules.append(f"⚠️ 互斥代碼: {row['互斥代碼']}")
            return " / ".join(rules) if rules else "-"
        
        df_courses['規則提醒'] = df_courses.apply(format_rules, axis=1)
        
        df_summary = all_sheets.get("學程規範總額", pd.DataFrame())
        return df_courses, df_summary
    except Exception as e:
        st.error(f"讀取失敗：{e}")
        st.stop()

# --- 載入邏輯 ---
DATA_FILE = "master_data.xlsx"
if os.path.exists(DATA_FILE):
    df_courses, df_summary = load_all_data(DATA_FILE)
else:
    st.error("找不到 master_data.xlsx，請確認檔案已上傳至 GitHub。")
    st.stop()

# 定義完整顯示的欄位清單
DISPLAY_COLS = ['科目類別', '課程代碼', '課程名稱', '學分數', '課程認抵範圍', '規則提醒']

# --- 2. 分頁標籤 ---
tab_browse, tab_search = st.tabs(["📂 按學程瀏覽", "🔍 依課程反查學程"])

# --- TAB 1: 按學程瀏覽 ---
with tab_browse:
    st.header("學分學程規範查詢")
    col1, col2, col3 = st.columns(3)
    with col1:
        college_list = ["全部"] + sorted(df_courses["學院"].dropna().unique().tolist())
        sel_col = st.selectbox("1. 選擇學院", college_list, key="b1")
    with col2:
        temp_p = df_courses if sel_col == "全部" else df_courses[df_courses["學院"] == sel_col]
        prog_list = sorted(temp_p["學程名稱"].dropna().unique().tolist())
        sel_prog = st.selectbox("2. 選擇學程", prog_list, key="b2")
    with col3:
        year_list = sorted(df_courses[df_courses["學程名稱"]==sel_prog]["適用年度"].unique(), reverse=True)
        sel_year = st.selectbox("3. 選擇年度", year_list, key="b3")

    summary = df_summary[(df_summary["學程名稱"] == sel_prog) & (df_summary["適用年度"] == sel_year)]
    if not summary.empty:
        s = summary.iloc[0]
        st.success(f"🎓 **應修學分：** 必修 {s.get('必修總學分',0)} / 選修 {s.get('選修總學分',0)} / 總計 {s.get('總計應修學分',0)} 學分")
        if pd.notna(s.get('備註 (模組要求)')):
            st.info(f"💡 **備註：** {s['備註 (模組要求)']}")
    
    res_df = df_courses[(df_courses["學程名稱"] == sel_prog) & (df_courses["適用年度"] == sel_year)]
    st.subheader("📋 完整應修科目表")
    st.dataframe(res_df[DISPLAY_COLS], use_container_width=True, hide_index=True)


# --- TAB 2: 依課程反查學程 (已簡化列表) ---
with tab_search:
    st.header("🔍 搜尋課程找學程")
    course_query = st.text_input("📝 請輸入課程名稱或代碼 (例如：微積分)", "", key="s1")
    
    if course_query:
        # 搜尋包含該課程的學程
        search_results = df_courses[
            df_courses["課程名稱"].str.contains(course_query, case=False, na=False) |
            df_courses["課程代碼"].str.contains(course_query, case=False, na=False)
        ].copy()
        
        if not search_results.empty:
            st.write(f"以下學程包含關鍵字「{course_query}」：")
            # --- 簡化後的列表：僅顯示學院、學程名稱、適用年度 ---
            mini_df = search_results[['學院', '學程名稱', '適用年度']].drop_duplicates()
            st.dataframe(mini_df, use_container_width=True, hide_index=True)
            
            st.divider()
            
            # 點選查看詳情
            st.subheader("📖 查看該學程完整資訊")
            # 建立選項清單
            options = mini_df.apply(lambda x: f"{x['學程名稱']} ({x['適用年度']})", axis=1).unique().tolist()
            selected_option = st.selectbox("請選擇上方搜尋結果中的學程：", ["--- 請選擇 ---"] + options, key="s2")
            
            if selected_option != "--- 請選擇 ---":
                p_name = selected_option.split(" (")[0]
                p_year = selected_option.split(" (")[1].replace(")", "")
                
                detail_df = df_courses[(df_courses["學程名稱"] == p_name) & (df_courses["適用年度"] == p_year)]
                detail_summary = df_summary[(df_summary["學程名稱"] == p_name) & (df_summary["適用年度"] == p_year)]
                
                with st.expander(f"📌 {selected_option} 完整應修科目表", expanded=True):
                    if not detail_summary.empty:
                        ds = detail_summary.iloc[0]
                        st.info(f"**應修學分：** 必修 {ds.get('必修總學分',0)} / 選修 {ds.get('選修總學分',0)} / 總計 {ds.get('總計應修學分',0)} 學分")
                    
                    # 顯示完整表格（包含所有詳細欄位與規則提醒）
                    st.dataframe(detail_df[DISPLAY_COLS], use_container_width=True, hide_index=True)
        else:
            st.warning("查無相關課程。")