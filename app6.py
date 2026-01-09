import streamlit as st
import pandas as pd
import os
import io

st.set_page_config(page_title="學分學程查詢系統", layout="wide")

# --- 1. 自定義 CSS (美化模組卡片樣式) ---
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
    </style>
    """, unsafe_allow_html=True)

# --- 2. 資料讀取與處理 ---
@st.cache_data
def load_all_data(file_path):
    try:
        all_sheets = pd.read_excel(file_path, sheet_name=None, dtype=str)
        df_courses = all_sheets["科目表"]
        # 轉學分為數字
        df_courses['學分數'] = pd.to_numeric(df_courses['學分數'], errors='coerce').fillna(0)
        
        # 格式化規則提醒 (僅保留互斥代碼，因選修群組已整合進模組標題)
        def format_rules(row):
            rules = []
            if pd.notna(row.get('互斥代碼')) and str(row['互斥代碼']).strip() != "":
                rules.append(f"⚠️ 互斥代碼: {row['互斥代碼']}")
            return " / ".join(rules) if rules else "-"
        
        df_courses['規則提醒'] = df_courses.apply(format_rules, axis=1)
        
        # 處理模組名稱：若空白則填入"一般科目"
        df_courses['模組名稱'] = df_courses.get('模組名稱', "一般科目").fillna("一般科目").replace("", "一般科目")
        
        df_summary = all_sheets.get("學程規範總額", pd.DataFrame())
        return df_courses, df_summary
    except Exception as e:
        st.error(f"讀取 Excel 失敗，請檢查格式：{e}")
        st.stop()

# 樣式函數：搜尋關鍵字標色
def highlight_keywords(val):
    query = st.session_state.get('last_query', "")
    if query and query.lower() in str(val).lower():
        return 'background-color: #FFF9C4; color: #E65100; font-weight: bold; border: 1px solid #FFCC80'
    return ''

# 樣式函數：認抵範圍關鍵字標色
def highlight_scope(val):
    target_keywords = ["通識", "系外", "跨系", "自由選修"]
    for key in target_keywords:
        if key in str(val):
            return 'background-color: #E8F5E9; color: #2E7D32; font-weight: bold'
    return ''

# 下載功能：轉檔為 Excel
def convert_df_to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='學程課表')
    return output.getvalue()

# --- 3. 核心顯示函數 (必選修 -> 模組卡片) ---
def display_styled_tables(df, is_search_mode=False):
    """依照 必/選修 -> 模組 結構顯示美化表格"""
    for cat in ["必修", "選修"]:
        cat_df = df[df["科目類別"] == cat]
        if not cat_df.empty:
            st.markdown(f"### 📍 {cat}課程清單")
            
            # 依照模組分組
            for mod in cat_df["模組名稱"].unique():
                mod_df = cat_df[cat_df["模組名稱"] == mod]
                total_credits = mod_df['學分數'].sum()
                
                # 建立美化卡片外殼
                st.markdown(f"""
                    <div class="module-card">
                        <div class="module-title">🔹 {mod}</div>
                        <div class="module-info">ℹ️ 此區塊共有 {len(mod_df)} 門課程，總計開放 {total_credits} 學分供選修</div>
                """, unsafe_allow_html=True)
                
                # 顯示表格
                cols = ['課程代碼', '課程名稱', '學分數', '課程認抵範圍', '規則提醒']
                styled_df = mod_df[cols].style \
                    .applymap(highlight_keywords, subset=['課程代碼', '課程名稱']) \
                    .applymap(highlight_scope, subset=['課程認抵範圍'])
                
                st.dataframe(styled_df, use_container_width=True, hide_index=True)
                st.markdown('</div>', unsafe_allow_html=True)
        elif not is_search_mode:
            st.write(f"（無{cat}科目相關資料）")

# --- 4. 主程式邏輯 ---
DATA_FILE = "master_data.xlsx"
if os.path.exists(DATA_FILE):
    df_courses, df_summary = load_all_data(DATA_FILE)
else:
    st.error("找不到 master_data.xlsx，請確認檔案已上傳。")
    st.stop()

tab_browse, tab_search = st.tabs(["📂 按學程瀏覽", "🔍 依課程反查學程"])

# --- TAB 1: 按學程瀏覽 ---
with tab_browse:
    st.header("學分學程規範查詢")
    
    # 建立連動選單
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

    # 取得當前學程資料
    prog_data = df_courses[(df_courses["學程名稱"] == sel_prog) & (df_courses["適用年度"] == sel_year)]
    
    # 標題與下載按鈕
    st.divider()
    head_col, down_col = st.columns([4, 1])
    with head_col:
        st.subheader(f"🎓 {sel_prog} ({sel_year}年度)")
    with down_col:
        excel_file = convert_df_to_excel(prog_data)
        st.download_button(label="📥 下載完整 Excel", data=excel_file, file_name=f"{sel_prog}_{sel_year}.xlsx", mime="application/vnd.ms-excel", use_container_width=True)

    # 顯示畢業門檻資訊
    summary = df_summary[(df_summary["學程名稱"] == sel_prog) & (df_summary["適用年度"] == sel_year)]
    if not summary.empty:
        s = summary.iloc[0]
        st.success(f"**畢業門檻：** 必修 {s.get('必修總學分',0)} / 選修 {s.get('選修總學分',0)} / 總計 {s.get('總計應修學分',0)} 學分")
        if pd.notna(s.get('備註 (模組要求)')): st.info(f"💡 **備註：** {s['備註 (模組要求)']}")

    # 呼叫分組表格顯示
    display_styled_tables(prog_data)


# --- TAB 2: 依課程反查學程 ---
with tab_search:
    st.header("🔍 搜尋課程找學程")
    
    # 清除搜尋函數
    def clear_search():
        st.session_state.s1 = ""
        st.session_state.last_query = ""
        if 's2' in st.session_state: st.session_state.s2 = "--- 請選擇 ---"

    col_search, col_clear = st.columns([4, 1])
    with col_search:
        course_query = st.text_input("📝 請輸入課程名稱或代碼關鍵字", key="s1")
        st.session_state.last_query = course_query # 同步至樣式函數
    with col_clear:
        st.markdown("<br>", unsafe_allow_html=True)
        st.button("🧹 清除條件", on_click=clear_search, use_container_width=True)

    if course_query:
        # 在全域搜尋
        search_res = df_courses[
            df_courses["課程名稱"].str.contains(course_query, case=False, na=False) |
            df_courses["課程代碼"].str.contains(course_query, case=False, na=False)
        ].copy()
        
        if not search_res.empty:
            st.write(f"以下學程包含關鍵字「{course_query}」：")
            # 簡化列表僅顯示 學院、學程、年度、課程代碼/名稱(用於對照)
            mini_df = search_res[['學院', '學程名稱', '適用年度', '課程代碼', '課程名稱']].drop_duplicates()
            styled_mini = mini_df.style.applymap(highlight_keywords, subset=['課程代碼', '課程名稱'])
            st.dataframe(styled_mini, use_container_width=True, hide_index=True)
            
            st.divider()
            
            # 點選詳細學程
            st.subheader("📖 查看完整學程資訊")
            prog_options = mini_df.apply(lambda x: f"{x['學程名稱']} ({x['適用年度']})", axis=1).unique().tolist()
            selected_prog = st.selectbox("請選擇上述搜尋結果中的學程：", ["--- 請選擇 ---"] + prog_options, key="s2")
            
            if selected_prog != "--- 請選擇 ---":
                p_name = selected_prog.split(" (")[0]
                p_year = selected_prog.split(" (")[1].replace(")", "")
                
                detail_data = df_courses[(df_courses["學程名稱"] == p_name) & (df_courses["適用年度"] == p_year)]
                detail_summary = df_summary[(df_summary["學程名稱"] == p_name) & (df_summary["適用年度"] == p_year)]
                
                with st.expander(f"📌 {selected_prog} 完整應修科目表", expanded=True):
                    if not detail_summary.empty:
                        ds = detail_summary.iloc[0]
                        st.info(f"**門檻：** 必修 {ds.get('必修總學分',0)} / 選修 {ds.get('選修總學分',0)} / 總計 {ds.get('總計應修學分',0)} 學分")
                    
                    display_styled_tables(detail_data, is_search_mode=True)
        else:
            st.warning("查無相關課程，請嘗試其他關鍵字。")