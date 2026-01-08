import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="學分學程查詢系統", layout="wide")

# --- 1. 資料讀取函數 ---
@st.cache_data
def load_all_data(file_path):
    df_courses = pd.read_excel(file_path, sheet_name="科目表", dtype=str)
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
    st.error(f"找不到 {DATA_FILE}")
    st.stop()

# --- 2. 側邊欄篩選器 ---
st.sidebar.header("🔍 系統篩選")
college_list = ["全部"] + sorted(df_courses["學院"].dropna().unique().tolist())
sel_college = st.sidebar.selectbox("1. 選擇學院", college_list)

if sel_college != "全部":
    temp_programs = df_courses[df_courses["學院"] == sel_college]
else:
    temp_programs = df_courses
program_list = sorted(temp_programs["學程名稱"].dropna().unique().tolist())
sel_program = st.sidebar.selectbox("2. 選擇學分學程", program_list)

year_list = sorted(df_courses[df_courses["學程名稱"]==sel_program]["適用年度"].unique(), reverse=True)
sel_year = st.sidebar.selectbox("3. 選擇適用年度", year_list)
search_query = st.sidebar.text_input("📝 搜尋課程名稱或代碼", "")

# --- 3. 執行過濾邏輯 ---
mask = (df_courses["學程名稱"] == sel_program) & (df_courses["適用年度"] == sel_year)
final_df = df_courses[mask].copy()

if search_query:
    final_df = final_df[
        final_df["課程名稱"].str.contains(search_query, case=False, na=False) |
        final_df["課程代碼"].str.contains(search_query, case=False, na=False)
    ]

# 讀取規則格式化函數
def format_rules(row):
    rules = []
    if pd.notna(row.get('選修群組')) and row['選修群組'] != "":
        rules.append(f"📦 {row['選修群組']} ({row.get('群組要求', '無要求')})")
    if pd.notna(row.get('互斥代碼')) and row['互斥代碼'] != "":
        rules.append(f"⚠️ 互斥代碼: {row['互斥代碼']}")
    return " / ".join(rules) if rules else "-"

final_df['規則提醒'] = final_df.apply(format_rules, axis=1)
available_cols = ['課程代碼', '課程名稱', '學分數', '課程認抵範圍', '規則提醒']

# --- 4. 畫面顯示 ---
st.title(f"🎓 {sel_program}")
st.caption(f"適用年度：{sel_year} 學年度")

# A. 學分儀表板
summary_row = df_summary[(df_summary["學程名稱"] == sel_program) & (df_summary["適用年度"] == sel_year)]
if not summary_row.empty:
    row = summary_row.iloc[0]
    c1, c2, c3 = st.columns(3)
    c1.metric("必修學分門檻", f"{row.get('必修總學分', '0')} 學分")
    c2.metric("選修學分門檻", f"{row.get('選修總學分', '0')} 學分")
    c3.metric("總計畢業要求", f"{row.get('總計應修學分', '0')} 學分")
    # 這裡顯示您問的「額外說明」
    if pd.notna(row.get('備註 (模組要求)')):
        st.info(f"💡 **學程額外說明：** {row['備註 (模組要求)']}")

st.divider()

# B. 分開表格顯示
required_df = final_df[final_df["科目類別"] == "必修"]
elective_df = final_df[final_df["科目類別"] == "選修"]

# 顯示必修表格
st.subheader("📌 必修科目")
if not required_df.empty:
    st.dataframe(required_df[available_cols], use_container_width=True, hide_index=True)
else:
    st.write("此條件下無必修科目")

st.markdown("<br>", unsafe_allow_html=True) # 增加間距

# 顯示選修表格
st.subheader("💡 選修科目")
if not elective_df.empty:
    st.dataframe(elective_df[available_cols], use_container_width=True, hide_index=True)
else:
    st.write("此條件下無選修科目")

# --- 5. 下載功能 ---
csv = final_df.to_csv(index=False).encode('utf-8-sig')
st.sidebar.download_button("📥 下載目前科目表", csv, f"{sel_program}.csv", "text/csv")