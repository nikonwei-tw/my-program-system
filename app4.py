import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="學分學程查詢系統", layout="wide")

# --- 1. 資料讀取 ---
@st.cache_data
def load_all_data(file_path):
    try:
        all_sheets = pd.read_excel(file_path, sheet_name=None, dtype=str)
        df_courses = all_sheets["科目表"]
        df_courses['學分數'] = pd.to_numeric(df_courses['學分數'], errors='coerce').fillna(0)
        df_summary = all_sheets.get("學程規範總額", pd.DataFrame())
        return df_courses, df_summary
    except Exception as e:
        st.error(f"讀取失敗：{e}")
        st.stop()

DATA_FILE = "master_data.xlsx"
if os.path.exists(DATA_FILE):
    df_courses, df_summary = load_all_data(DATA_FILE)
else:
    st.error("找不到 master_data.xlsx")
    st.stop()

# --- 2. 分頁標籤 ---
tab_browse, tab_search = st.tabs(["📂 按學程瀏覽", "🔍 依課程反查學程"])

# --- TAB 1: 按學程瀏覽 ---
with tab_browse:
    st.header("學分學程規範查詢")
    col1, col2, col3 = st.columns(3)
    with col1:
        college_list = ["全部"] + sorted(df_courses["學院"].dropna().unique().tolist())
        sel_col = st.selectbox("選擇學院", college_list)
    with col2:
        temp_p = df_courses if sel_col == "全部" else df_courses[df_courses["學院"] == sel_col]
        prog_list = sorted(temp_p["學程名稱"].dropna().unique().tolist())
        sel_prog = st.selectbox("選擇學程", prog_list)
    with col3:
        year_list = sorted(df_courses[df_courses["學程名稱"]==sel_prog]["適用年度"].unique(), reverse=True)
        sel_year = st.selectbox("選擇年度", year_list)

    # 顯示畢業門檻與列表
    summary = df_summary[(df_summary["學程名稱"] == sel_prog) & (df_summary["適用年度"] == sel_year)]
    if not summary.empty:
        s = summary.iloc[0]
        st.success(f"🎓 **畢業門檻：** 必修 {s.get('必修總學分',0)} / 選修 {s.get('選修總學分',0)} / 總計 {s.get('總計應修學分',0)} 學分")
    
    res_df = df_courses[(df_courses["學程名稱"] == sel_prog) & (df_courses["適用年度"] == sel_year)]
    st.subheader("📋 完整應修科目表")
    st.dataframe(res_df[['科目類別', '課程代碼', '課程名稱', '學分數', '課程認抵範圍']], use_container_width=True, hide_index=True)


# --- TAB 2: 依課程反查學程 (關鍵功能更新) ---
with tab_search:
    st.header("🔍 搜尋課程找學程")
    course_query = st.text_input("📝 請輸入課程名稱或代碼 (例如：微積分)", "")
    
    if course_query:
        # 1. 搜尋包含該課程的學程
        search_results = df_courses[
            df_courses["課程名稱"].str.contains(course_query, case=False, na=False) |
            df_courses["課程代碼"].str.contains(course_query, case=False, na=False)
        ].copy()
        
        if not search_results.empty:
            # 2. 顯示簡易搜尋結果表格
            st.write(f"以下學程包含「{course_query}」：")
            mini_df = search_results[['學院', '學程名稱', '適用年度', '科目類別']].drop_duplicates()
            st.dataframe(mini_df, use_container_width=True, hide_index=True)
            
            st.divider()
            
            # 3. 提供下拉選單讓使用者「點選」感興趣的學程詳情
            st.subheader("📖 查看完整學程資訊")
            # 建立選項清單： "學程名稱 (年度)"
            options = mini_df.apply(lambda x: f"{x['學程名稱']} ({x['適用年度']})", axis=1).tolist()
            selected_option = st.selectbox("請選擇上方搜尋結果中的學程，以查看完整應修科目：", ["--- 請選擇 ---"] + options)
            
            if selected_option != "--- 請選擇 ---":
                # 解析出名稱與年度
                p_name = selected_option.split(" (")[0]
                p_year = selected_option.split(" (")[1].replace(")", "")
                
                # 抓取該學程的完整資料
                detail_df = df_courses[(df_courses["學程名稱"] == p_name) & (df_courses["適用年度"] == p_year)]
                detail_summary = df_summary[(df_summary["學程名稱"] == p_name) & (df_summary["適用年度"] == p_year)]
                
                # 顯示詳細資訊卡片
                with st.expander(f"📌 {selected_option} 完整應修科目表", expanded=True):
                    if not detail_summary.empty:
                        ds = detail_summary.iloc[0]
                        st.info(f"**畢業門檻：** 必修 {ds.get('必修總學分',0)} / 選修 {ds.get('選修總學分',0)} / 總計 {ds.get('總計應修學分',0)} 學分")
                    
                    # 必選修分表顯示
                    st.write("**必修科目**")
                    st.table(detail_df[detail_df["科目類別"]=="必修"][['課程代碼', '課程名稱', '學分數']])
                    st.write("**選修科目**")
                    st.table(detail_df[detail_df["科目類別"]=="選修"][['課程代碼', '課程名稱', '學分數']])
        else:
            st.warning("查無相關課程。")