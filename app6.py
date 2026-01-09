import streamlit as st
import pandas as pd
import os
import io

st.set_page_config(page_title="學分學程查詢系統", layout="wide")

# --- 1. 自定義 CSS (美化模組樣式) ---
st.markdown("""
    <style>
    .module-card {
        background-color: #f8f9fa;
        border-radius: 10px;
        padding: 15px;
        margin-bottom: 15px;
        border-left: 5px solid #007bff;
        box-shadow: 2px 2px 5px rgba(0,0,0,0.05);
    }
    .module-title {
        color: #0056b3;
        font-weight: bold;
        font-size: 1.1rem;
        margin-bottom: 10px;
    }
    </style>
    """, unsafe_allow_html=True)

# --- 2. 資料讀取 ---
@st.cache_data
def load_all_data(file_path):
    try:
        all_sheets = pd.read_excel(file_path, sheet_name=None, dtype=str)
        df_courses = all_sheets["科目表"]
        df_courses['學分數'] = pd.to_numeric(df_courses['學分數'], errors='coerce').fillna(0)
        
        def format_rules(row):
            rules = []
            if pd.notna(row.get('選修群組')) and str(row['選修群組']).strip() != "":
                rules.append(f"📦 {row['選修群組']} ({row.get('群組要求', '無要求')})")
            if pd.notna(row.get('互斥代碼')) and str(row['互斥代碼']).strip() != "":
                rules.append(f"⚠️ 互斥代碼: {row['互斥代碼']}")
            return " / ".join(rules) if rules else "-"
        
        df_courses['規則提醒'] = df_courses.apply(format_rules, axis=1)
        df_courses['模組名稱'] = df_courses.get('模組名稱', "一般科目").fillna("一般科目").replace("", "一般科目")
        
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
    for key in ["通識", "系外", "跨系", "自由選修"]:
        if key in str(val): return 'background-color: #E8F5E9; color: #2E7D32; font-weight: bold'
    return ''

# --- 3. 下載功能邏輯 ---
def convert_df_to_excel(df, prog_name):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='學程課表')
    return output.getvalue()

# --- 4. 顯示美化表格函數 ---
def display_styled_tables(df):
    for cat in ["必修", "選修"]:
        cat_df = df[df["科目類別"] == cat]
        if not cat_df.empty:
            st.markdown(f"#### 📍 {cat}科目清單")
            for mod in cat_df["模組名稱"].unique():
                mod_df = cat_df[cat_df["模組名稱"] == mod]
                # 使用自定義 CSS 容器
                st.markdown(f'<div class="module-card"><div class="module-title">🔹 {mod}</div>', unsafe_allow_html=True)
                
                cols = ['課程代碼', '課程名稱', '學分數', '課程認抵範圍', '規則提醒']
                styled_df = mod_df[cols].style \
                    .applymap(highlight_keywords, subset=['課程代碼', '課程名稱']) \
                    .applymap(highlight_scope, subset=['課程認抵範圍'])
                
                st.dataframe(styled_df, use_container_width=True, hide_index=True)
                st.markdown('</div>', unsafe_allow_html=True)

# --- 5. 主程式 ---
DATA_FILE = "master_data.xlsx"
if os.path.exists(DATA_FILE):
    df_courses, df_summary = load_all_data(DATA_FILE)
else:
    st.stop()

tab_browse, tab_search = st.tabs(["📂 按學程瀏覽", "🔍 依課程反查學程"])

with tab_browse:
    col_b1, col_b2, col_b3, col_reset = st.columns([2, 2, 2, 1])
    with col_b1: sel_col = st.selectbox("選擇學院", ["全部"] + sorted(df_courses["學院"].dropna().unique().tolist()), key="b1")
    with col_b2: sel_prog = st.selectbox("選擇學程", sorted(df_courses[df_courses["學院"]==sel_col]["學程名稱"].unique()) if sel_col!="全部" else sorted(df_courses["學程名稱"].unique()), key="b2")
    with col_b3: sel_year = st.selectbox("選擇年度", sorted(df_courses[df_courses["學程名稱"]==sel_prog]["適用年度"].unique(), reverse=True), key="b3")
    with col_reset: 
        st.markdown("<br>", unsafe_allow_html=True)
        st.button("🔄 重設", on_click=lambda: st.session_state.update(b1="全部"), use_container_width=True)

    # 顯示內容
    prog_data = df_courses[(df_courses["學程名稱"] == sel_prog) & (df_courses["適用年度"] == sel_year)]
    
    # 下載按鈕 (放在標題旁邊)
    header_col, download_col = st.columns([4, 1])
    with header_col: st.title(f"🎓 {sel_prog}")
    with download_col:
        st.markdown("<br>", unsafe_allow_html=True)
        excel_data = convert_df_to_excel(prog_data, sel_prog)
        st.download_button(label="📥 下載完整 Excel 課表", data=excel_data, file_name=f"{sel_prog}_{sel_year}.xlsx", mime="application/vnd.ms-excel", use_container_width=True)

    display_styled_tables(prog_data)

# TAB 2 邏輯同樣呼叫 display_styled_tables(detail_df) 即可，此處為節省空間略，邏輯一致