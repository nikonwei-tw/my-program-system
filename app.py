import streamlit as st
import pandas as pd

# 網頁基礎設定
st.set_page_config(page_title="學分學程管理系統", layout="wide")
st.title("🎓 學分學程應修科目查詢系統")

# --- 檔案上傳 ---
uploaded_file = st.sidebar.file_uploader("1. 上傳學程 Excel 規範檔", type=["xlsx"])

if uploaded_file:
    # 讀取資料 (強制將代碼轉為字串避免遺漏字首 0)
    df = pd.read_excel(uploaded_file, dtype={'學程代碼': str, '課程代碼': str, '適用年度': str})
    
    # --- 側邊欄篩選邏輯 ---
    st.sidebar.header("2. 查詢篩選條件")
    
    colleges = ["全部"] + sorted(df["學院"].unique().tolist())
    sel_college = st.sidebar.selectbox("選擇學院", colleges)
    
    # 動態過濾學程清單
    filtered_programs = df if sel_college == "全部" else df[df["學院"] == sel_college]
    programs = ["全部"] + sorted(filtered_programs["學程名稱"].unique().tolist())
    sel_program = st.sidebar.selectbox("選擇學分學程", programs)
    
    years = ["全部"] + sorted(df["適用年度"].unique().tolist(), reverse=True)
    sel_year = st.sidebar.selectbox("選擇適用年度", years)

    # --- 執行篩選 ---
    final_df = df.copy()
    if sel_college != "全部":
        final_df = final_df[final_df["學院"] == sel_college]
    if sel_program != "全部":
        final_df = final_df[final_df["學程名稱"] == sel_program]
    if sel_year != "全部":
        final_df = final_df[final_df["適用年度"] == sel_year]

    # --- 顯示結果 ---
    st.subheader(f"📊 {sel_program if sel_program != '全部' else '所有學程'} - 應修科目表")
    
    if not final_df.empty:
        # 增加提示邏輯：美化顯示群組與互斥資訊
        def highlight_rules(row):
            rules = []
            if pd.notna(row['選修群組']):
                rules.append(f"📦 {row['選修群組']} ({row['群組要求']})")
            if pd.notna(row['互斥代碼']):
                rules.append(f"⚠️ 互斥代碼: {row['互斥代碼']}")
            return " / ".join(rules) if rules else "無特殊限制"

        final_df['特殊規則說明'] = final_df.apply(highlight_rules, axis=1)

        # 重新排版表格顯示順序
        display_cols = ['科目類別', '課程代碼', '課程名稱', '學分數', '課程認抵範圍', '特殊規則說明']
        st.dataframe(final_df[display_cols], use_container_width=True, hide_index=True)

        # 統計資訊區
        st.info(f"💡 本學程共有 {len(final_df)} 門課程。提醒：請注意選修群組之達標要求。")
        
        # 下載按鈕
        csv = final_df.to_csv(index=False).encode('utf-8-sig')
        st.download_button("📥 下載篩選結果", csv, "program_requirements.csv", "text/csv")
    else:
        st.warning("查無對應資料，請檢查 Excel 內容或篩選條件。")
else:
    st.info("👋 請在左側上傳學程規範 Excel 檔案。系統將自動解析複雜的選修與互斥邏輯。")

# 在 app.py 中加入此段 (建議放在 sidebar)
with st.sidebar:
    st.divider()
    st.write("### 範本下載")
    # 假設您在 GitHub 也有上傳一個名為 template.xlsx 的檔案
    with open("example_data.xlsx", "rb") as file:
        st.download_button(
            label="📥 下載 Excel 填寫範本",
            data=file,
            file_name="學分學程填寫範本.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )