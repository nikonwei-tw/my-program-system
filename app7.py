import streamlit as st
import pandas as pd
import os
import io

st.set_page_config(page_title="學分學程查詢系統", layout="wide")

# --- 1. 自定義 CSS ---
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
    .stProgress > div > div > div > div {
        background-color: #28a745;
    }
    </style>
    """, unsafe_allow_html=True)

# --- 2. 資料讀取與處理 ---
@st.cache_data
def load_all_data(file_path):
    try:
        all_sheets = pd.read_excel(file_path, sheet_name=None, dtype=str)
        df_courses = all_sheets["科目表"]
        df_courses['學分數'] = pd.to_numeric(df_courses['學分數'], errors='coerce').fillna(0)
        
        def format_rules(row):
            rules = []
            if pd.notna(row.get('互斥代碼')) and str(row['互斥代碼']).strip() != "":
                rules.append(f"⚠️ 互斥代碼: {row['互斥代碼']}")
            return " / ".join(rules) if rules else "-"
        
        df_courses['規則提醒'] = df_courses.apply(format_rules, axis=1)
        df_courses['模組名稱'] = df_courses.get('模組名稱', "一般科目").fillna("一般科目").replace("", "一般科目")
        df_summary = all_sheets.get("學程規範總額", pd.DataFrame())
        return df_courses, df_summary
    except Exception as e:
        st.error(f"讀取 Excel 失敗：{e}")
        st.stop()

# 樣式函數與下載功能 (維持原樣)
def highlight_keywords(val):
    query = st.session_state.get('last_query', "")
    if query and query.lower() in str(val).lower():
        return 'background-color: #FFF9C4; color: #E65100; font-weight: bold'
    return ''

def highlight_scope(val):
    for key in ["通識", "系外", "跨系", "自由選修"]:
        if key in str(val): return 'background-color: #E8F5E9; color: #2E7D32; font-weight: bold'
    return ''

def convert_df_to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='學程課表')
    return output.getvalue()

def display_styled_tables(df, is_search_mode=False):
    for cat in ["必修", "選修"]:
        cat_df = df[df["科目類別"] == cat]
        if not cat_df.empty:
            st.markdown(f"### 📍 {cat}課程清單")
            for mod in cat_df["模組名稱"].unique():
                mod_df = cat_df[cat_df["模組名稱"] == mod]
                st.markdown(f'<div class="module-card"><div class="module-title">🔹 {mod}</div><div class="module-info">ℹ️ 共有 {len(mod_df)} 門課，總開放 {mod_df["學分數"].sum()} 學分</div>', unsafe_allow_html=True)
                cols = ['課程代碼', '課程名稱', '學分數', '課程認抵範圍', '規則提醒']
                st.dataframe(mod_df[cols].style.applymap(highlight_keywords, subset=['課程代碼', '課程名稱']).applymap(highlight_scope, subset=['課程認抵範圍']), use_container_width=True, hide_index=True)
                st.markdown('</div>', unsafe_allow_html=True)

# --- 3. 進度比對核心函數 ---
def display_interactive_progress(df, summary_row):
    st.subheader("📊 畢業進度即時試算")
    
    # 初始化勾選狀態
    if 'checked_courses' not in st.session_state:
        st.session_state.checked_courses = set()

    plot_df = df.copy()
    plot_df.insert(0, '已修過', plot_df['課程代碼'].apply(lambda x: x in st.session_state.checked_courses))

    # 互動式編輯器
    edited_df = st.data_editor(
        plot_df[['已修過', '科目類別', '模組名稱', '課程代碼', '課程名稱', '學分數']],
        key="progress_editor",
        hide_index=True,
        use_container_width=True,
        disabled=['科目類別', '模組名稱', '課程代碼', '課程名稱', '學分數']
    )

    # 同步勾選狀態
    st.session_state.checked_courses = set(edited_df[edited_df['已修過'] == True]['課程代碼'].tolist())

    # 計算進度
    done_req = edited_df[(edited_df['已修過']) & (edited_df['科目類別'] == '必修')]['學分數'].sum()
    done_elec = edited_df[(edited_df['已修過']) & (edited_df['科目類別'] == '選修')]['學分數'].sum()
    
    g_req = pd.to_numeric(summary_row.get('必修總學分', 0), errors='coerce')
    g_elec = pd.to_numeric(summary_row.get('選修總學分', 0), errors='coerce')
    g_total = pd.to_numeric(summary_row.get('總計應修學分', 0), errors='coerce')
    
    cur_total = done_req + done_elec
    progress_pct = min(cur_total / g_total, 1.0) if g_total > 0 else 0.0

    # 儀表板顯示
    c1, c2, c3 = st.columns(3)
    c1.metric("必修進度", f"{int(done_req)} / {int(g_req)} 學分")
    c2.metric("選修進度", f"{int(done_elec)} / {int(g_elec)} 學分")
    c3.metric("總達成率", f"{int(progress_pct*100)}%")
    
    st.progress(progress_pct)
    if cur_total >= g_total:
        st.balloons()
        st.success("🎉 恭喜！您已達到畢業學分門檻！")
    else:
        st.warning(f"🚩 尚差 {int(g_total - cur_total)} 學分即可取得學程證書。")

# --- 4. 主程式 ---
DATA_FILE = "master_data.xlsx"
if os.path.exists(DATA_FILE):
    df_courses, df_summary = load_all_data(DATA_FILE)
else:
    st.stop()

tab_browse, tab_search, tab_audit = st.tabs(["📂 按學程瀏覽", "🔍 依課程反查學程", "🎓 畢業進度試算"])

# --- TAB 1 & 2 內容與原本 app6.py 一致 (省略以節省空間) ---
with tab_browse:
    # ... 原有代碼 ...
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
    with head_col: st.subheader(f"🎓 {sel_prog} ({sel_year}年度)")
    with down_col:
        excel_file = convert_df_to_excel(prog_data)
        st.download_button(label="📥 下載完整 Excel", data=excel_file, file_name=f"{sel_prog}_{sel_year}.xlsx", mime="application/vnd.ms-excel", use_container_width=True)
    summary = df_summary[(df_summary["學程名稱"] == sel_prog) & (df_summary["適用年度"] == sel_year)]
    if not summary.empty:
        s = summary.iloc[0]
        st.success(f"**畢業門檻：** 必修 {s.get('必修總學分',0)} / 選修 {s.get('選修總學分',0)} / 總計 {s.get('總計應修學分',0)} 學分")
        if pd.notna(s.get('備註 (模組要求)')): st.info(f"💡 **備註：** {s['備註 (模組要求)']}")
    display_styled_tables(prog_data)

with tab_search:
    # ... 原有代碼 ...
    st.header("🔍 搜尋課程找學程")
    col_search, col_clear = st.columns([4, 1])
    with col_search:
        course_query = st.text_input("📝 請輸入課程名稱或代碼關鍵字", key="s1")
        st.session_state.last_query = course_query
    with col_clear:
        st.markdown("<br>", unsafe_allow_html=True)
        st.button("🧹 清除條件", on_click=lambda: st.session_state.update(s1="", last_query=""), use_container_width=True)
    if course_query:
        search_res = df_courses[df_courses["課程名稱"].str.contains(course_query, case=False, na=False) | df_courses["課程代碼"].str.contains(course_query, case=False, na=False)].copy()
        if not search_res.empty:
            mini_df = search_res[['學院', '學程名稱', '適用年度', '課程代碼', '課程名稱']].drop_duplicates()
            st.dataframe(mini_df.style.applymap(highlight_keywords, subset=['課程代碼', '課程名稱']), use_container_width=True, hide_index=True)
            st.divider()
            prog_options = mini_df.apply(lambda x: f"{x['學程名稱']} ({x['適用年度']})", axis=1).unique().tolist()
            selected_prog = st.selectbox("請選擇上述搜尋結果中的學程：", ["--- 請選擇 ---"] + prog_options, key="s2")
            if selected_prog != "--- 請選擇 ---":
                p_n = selected_prog.split(" (")[0]
                p_y = selected_prog.split(" (")[1].replace(")", "")
                d_data = df_courses[(df_courses["學程名稱"] == p_n) & (df_courses["適用年度"] == p_y)]
                display_styled_tables(d_data, is_search_mode=True)
        else: st.warning("查與相關課程。")

# --- TAB 3: 畢業進度試算 ---
with tab_audit:
    st.header("🎓 學程修習進度自動試算")
    st.markdown("請選擇學程並上傳成績單，系統將自動比對您的修課進度。")
    
    # 選項設定
    c1, c2 = st.columns(2)
    with c1:
        aud_prog = st.selectbox("1. 選擇要試算的學程", sorted(df_courses["學程名稱"].unique()), key="aud_p")
    with c2:
        aud_year = st.selectbox("2. 選擇適用年度", sorted(df_courses[df_courses["學程名稱"]==aud_prog]["適用年度"].unique(), reverse=True), key="aud_y")

    # 成績單上傳區
    uploaded_file = st.file_uploader("📥 上傳成績單 (Excel 格式，需含「課程代碼」欄位)", type=["xlsx"])

    # 取得基礎資料
    aud_data = df_courses[(df_courses["學程名稱"] == aud_prog) & (df_courses["適用年度"] == aud_year)].copy()
    aud_sum = df_summary[(df_summary["學程名稱"] == aud_prog) & (df_summary["適用年度"] == aud_year)]

    # 處理自動比對邏輯
    if uploaded_file:
        try:
            student_df = pd.read_excel(uploaded_file, dtype=str)
            if "課程代碼" in student_df.columns:
                passed_codes = set(student_df["課程代碼"].str.strip().tolist())
                # 更新狀態
                st.session_state.checked_courses = passed_codes
                st.success(f"✅ 比對完成！偵測到成績單中包含本學程的相關課程。")
            else:
                st.error("❌ 成績單格式錯誤：找不到「課程代碼」欄位。")
        except Exception as e:
            st.error(f"讀取檔案失敗：{e}")

    if not aud_sum.empty:
        st.divider()
        display_interactive_progress(aud_data, aud_sum.iloc[0])
    else:
        st.error("找不到該學程的規範門檻。")