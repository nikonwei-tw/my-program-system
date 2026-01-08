import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="學分學程查詢系統", layout="wide")

# --- 1. 資料讀取函數 ---
@st.cache_data
def load_all_data(file_path):
    # 讀取科目表
    df_courses = pd.read_excel(file_path, sheet_name="科目表", dtype=str)
    df_courses['學分數'] = pd.to_numeric(df_courses['學分數'], errors='coerce')
    
    # 讀取總額規範表
    try:
        df_summary = pd.read_excel(file_path, sheet_name="學程規範總額", dtype=str)
    except:
        df_summary = pd.DataFrame() # 若無此分頁則回傳空表
        
    return df_courses, df_summary

# --- 載入邏輯 ---
DATA_FILE = "master_data.xlsx"

if os.path.exists(DATA_FILE):
    df_courses, df_summary = load_all_data(DATA_FILE)
    st.sidebar.success("✅ 資料庫連線正常")
else:
    st.error(f"找不到 {DATA_FILE}，請確認檔案已上傳至 GitHub。")
    st.stop()

# --- 2. 側邊欄篩選 ---
st.sidebar.header("🔍 查詢篩選")
sel_program = st.sidebar.selectbox("1. 選擇學分學程", sorted(df_courses["學程名稱"].unique()))
sel_year = st.sidebar.selectbox("2. 選擇適用年度", sorted(df_courses[df_courses["學程名稱"]==sel_program]["適用年度"].unique(), reverse=True))

# --- 3. 篩選資料 ---
course_result = df_courses[(df_courses["學程名稱"] == sel_program) & (df_courses["適用年度"] == sel_year)]
summary_result = df_summary[(df_summary["學程名稱"] == sel_program) & (df_summary["適用年度"] == sel_year)]

# --- 4. 顯示結果 ---
st.title(f"🎓 {sel_program} ({sel_year}學年度)")

# --- A. 顯示學分總量指標 (Dashboard) ---
if not summary_result.empty:
    row = summary_result.iloc[0]
    c1, c2, c3 = st.columns(3)
    c1.metric("📌 必修學分", f"{row.get('必修總學分', '0')} 學分")
    c2.metric("📝 選修學分", f"{row.get('選修總學分', '0')} 學分")
    c3.metric("🏆 總計應修", f"{row.get('總計應修學分', '0')} 學分")
    
    if pd.notna(row.get('備註 (模組要求)')):
        st.warning(f"**特別要求：** {row['備註 (模組要求)']}")
else:
    st.info("⚠️ 尚未建立此學程的總額規範數據。")

st.divider()

# --- B. 分類顯示科目清單 ---
tab1, tab2 = st.tabs(["📋 必修科目", "💡 選修科目 (含模組)"])

with tab1:
    req_df = course_result[course_result["科目類別"] == "必修"]
    st.dataframe(req_df[['課程代碼', '課程名稱', '學分數', '課程認抵範圍']], use_container_width=True, hide_index=True)

with tab2:
    elec_df = course_result[course_result["科目類別"] == "選修"]
    if "模組名稱" in elec_df.columns:
        # 如果有模組，按模組分組顯示
        for module, group in elec_df.groupby("模組名稱"):
            module_name = module if pd.notna(module) else "一般選修"
            with st.expander(f"🔹 {module_name} (點擊展開)"):
                st.table(group[['課程代碼', '課程名稱', '學分數', '課程認抵範圍']])
    else:
        st.dataframe(elec_df[['課程代碼', '課程名稱', '學分數', '課程認抵範圍']], use_container_width=True, hide_index=True)

# --- 5. 下載功能 ---
csv = course_result.to_csv(index=False).encode('utf-8-sig')
st.download_button("📥 下載完整科目表", csv, f"{sel_program}_{sel_year}.csv", "text/csv")