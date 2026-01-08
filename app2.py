import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="學分學程查詢系統", layout="wide")

# --- 1. 資料讀取函數 ---
@st.cache_data
def load_all_data(file_path):
    df_courses = pd.read_excel(file_path, sheet_name="科目表", dtype=str)
    # 轉換學分為數字以便統計，轉換失敗則設為 0
    df_courses['學分數'] = pd.to_numeric(df_courses['學分數'], errors='coerce').fillna(0)
    
    try:
        df_summary = pd.read_excel(file_path, sheet_name="學程規範總額", dtype=str)
    except:
        df_summary = pd.DataFrame()
        
    return df_courses, df_summary

# --- 載入邏輯 ---
DATA_FILE = "master_data.xlsx"

if os.path.exists(DATA_FILE):
    df_courses, df_summary = load_all_data(DATA_FILE)
else:
    st.error(f"找不到 {DATA_FILE}，請確認檔案已上傳至 GitHub。")
    st.stop()

# --- 2. 側邊欄篩選器 (Sidebar) ---
st.sidebar.header("🔍 系統篩選")

# A. 學院篩選
college_list = ["全部"] + sorted(df_courses["學院"].dropna().unique().tolist())
sel_college = st.sidebar.selectbox("1. 選擇學院", college_list)

# B. 學程篩選 (連動學院)
if sel_college != "全部":
    temp_programs = df_courses[df_courses["學院"] == sel_college]
else:
    temp_programs = df_courses
program_list = sorted(temp_programs["學程名稱"].dropna().unique().tolist())
sel_program = st.sidebar.selectbox("2. 選擇學分學程", program_list)

# C. 年度篩選 (連動學程)
year_list = sorted(df_courses[df_courses["學程名稱"]==sel_program]["適用年度"].unique(), reverse=True)
sel_year = st.sidebar.selectbox("3. 選擇適用年度", year_list)

# D. 關鍵字搜尋
search_query = st.sidebar.text_input("📝 搜尋課程名稱或代碼", "")

# --- 3. 執行過濾邏輯 ---
# 基礎過濾：學程 + 年度
mask = (df_courses["學程名稱"] == sel_program) & (df_courses["適用年度"] == sel_year)
final_df = df_courses[mask].copy()

# 關鍵字過濾
if search_query:
    final_df = final_df[
        final_df["課程名稱"].str.contains(search_query, case=False, na=False) |
        final_df["課程代碼"].str.contains(search_query, case=False, na=False)
    ]

# 抓取該學程的總額規範
summary_row = df_summary[
    (df_summary["學程名稱"] == sel_program) & (df_summary["適用年度"] == sel_year)
]

# --- 4. 畫面顯示 ---
st.title(f"🎓 {sel_program}")
st.caption(f"適用年度：{sel_year} 學年度 ｜ 所屬學院：{df_courses[df_courses['學程名稱']==sel_program]['學院'].iloc[0]}")

# --- A. 顯示學分總量指標 (Dashboard) ---
st.markdown("### 📌 畢業學分要求")
if not summary_row.empty:
    row = summary_row.iloc[0]
    c1, c2, c3 = st.columns(3)
    c1.metric("必修學分", f"{row.get('必修總學分', '0')} 學分")
    c2.metric("選修學分", f"{row.get('選修總學分', '0')} 學分")
    c3.metric("總計應修", f"{row.get('總計應修學分', '0')} 學分")
    if pd.notna(row.get('備註 (模組要求)')):
        st.info(f"💡 **額外說明：** {row['備註 (模組要求)']}")
else:
    st.warning("⚠️ 尚未設定此學程的學分總量規範。")

st.divider()

# --- B. 顯示整合後的科目清單 ---
st.markdown("### 📋 應修科目清單")

if not final_df.empty:
    # 格式化規則提醒
    def format_rules(row):
        rules = []
        if pd.notna(row.get('選修群組')) and row['選修群組'] != "":
            rules.append(f"📦 {row['選修群組']} ({row.get('群組要求', '無要求')})")
        if pd.notna(row.get('互斥代碼')) and row['互斥代碼'] != "":
            rules.append(f"⚠️ 互斥代碼: {row['互斥代碼']}")
        return " / ".join(rules) if rules else "-"

    final_df['特殊規則說明'] = final_df.apply(format_rules, axis=1)

    # 欄位排序
    col_order = ['科目類別', '模組名稱', '課程代碼', '課程名稱', '學分數', '課程認抵範圍', '特殊規則說明']
    # 過濾掉不存在的欄位
    available_cols = [c for c in col_order if c in final_df.columns]
    
    # 顯示表格 (使用 dataframe 讓使用者可以自行點擊標題排序)
    st.dataframe(
        final_df[available_cols], 
        use_container_width=True, 
        hide_index=True
    )
    
    st.caption(f"共顯示 {len(final_df)} 門符合條件的課程。")
else:
    st.error("查無資料，請確認搜尋關鍵字或篩選條件。")

# --- 5. 下載功能 ---
csv = final_df.to_csv(index=False).encode('utf-8-sig')
st.download_button("📥 下載目前的查詢結果", csv, f"{sel_program}_list.csv", "text/csv")