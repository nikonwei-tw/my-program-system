import streamlit as st
import pandas as pd
import os
import io

# --- 0. 頁面設定 ---
st.set_page_config(page_title="學分學程查詢與比對系統", layout="wide")

# --- 1. 自定義 CSS (增加 Tab 3 模組區隔樣式) ---
st.markdown("""
    <style>
    .module-card {
        background-color: #f8f9fa; border-radius: 10px; padding: 18px;
        margin-bottom: 20px; border-left: 6px solid #007bff;
        box-shadow: 2px 2px 8px rgba(0,0,0,0.08);
    }
    .module-title { color: #0056b3; font-weight: bold; font-size: 1.15rem; margin-bottom: 5px; }
    .prog-container {
        background-color: #ffffff; border: 1px solid #e0e0e0;
        border-radius: 10px; padding: 20px; margin-bottom: 25px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
    }
    .audit-module-header {
        background-color: #e9ecef; padding: 8px 15px;
        border-radius: 5px; font-weight: bold; color: #495057;
        margin-top: 15px; margin-bottom: 10px;
        border-left: 4px solid #6c757d;
    }
    .note-box {
        background-color: #fff3cd; border-left: 5px solid #ffc107;
        padding: 12px; margin: 10px 0; font-size: 0.95rem; color: #856404;
    }
    </style>
    """, unsafe_allow_html=True)

# --- 2. 資料讀取與處理 (包含標題空格清理與名稱修正) ---
@st.cache_data
def load_all_data(file_path):
    try:
        all_sheets = pd.read_excel(file_path, sheet_name=None, dtype=str)
        
        # A. 科目表
        df_courses = all_sheets["科目表"]
        df_courses.columns = df_courses.columns.str.strip()
        df_courses['學分數'] = pd.to_numeric(df_courses['學分數'], errors='coerce').fillna(0)
        df_courses['模組名稱'] = df_courses.get('模組名稱', "一般科目").fillna("一般科目")
        
        # 修正名稱對照 (如 K860 -> 教育大數據)
        prog_map = df_courses[['學程代碼', '學程名稱']].drop_duplicates().set_index('學程代碼')['學程名稱'].to_dict()
        
        # B. 規範總額
        df_summary = all_sheets.get("學程規範總額", pd.DataFrame())
        df_summary.columns = df_summary.columns.str.strip()
        if not df_summary.empty and '學程代碼' in df_summary.columns:
            df_summary['學程名稱'] = df_summary['學程代碼'].map(prog_map).fillna(df_summary.get('學程名稱', '未知學程'))
            for col in ['必修總學分', '選修總學分', '總計應修學分']:
                if col in df_summary.columns:
                    df_summary[col] = pd.to_numeric(df_summary[col], errors='coerce').fillna(0.0)
        
        return df_courses, df_summary, prog_map
    except Exception as e:
        st.error(f"讀取失敗：{e}"); st.stop()

# --- 3. 核心功能函數 ---
def reset_filters():
    st.session_state.b1 = "全部"

def check_passing(grade):
    g = str(grade).strip().upper()
    if g in ["通過", "及格", "P", "PASS"]: return True
    try: return float(g) >= 60
    except: return False

def get_template_excel():
    """更新下載範本：加入單位代碼與範例"""
    df_temp = pd.DataFrame(columns=['課程代碼', '課程名稱', '學分數', '開課單位代碼', '成績'])
    df_temp.loc[0] = ['1001', '程式設計範例', '3', 'D51', '85']
    df_temp.loc[1] = ['1002', '微積分範例', '3', 'D89', '及格']
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_temp.to_excel(writer, index=False)
    return output.getvalue()

def check_course_completion(req_row, passed_df):
    """認抵比對 (支援多單位 D51, D89)"""
    matches = passed_df[passed_df['課程代碼'].str.strip() == str(req_row['課程代碼']).strip()]
    if matches.empty: return False
    allowed_val = str(req_row.get('認抵單位代碼', 'ANY')).strip().upper()
    if allowed_val in ["ANY", "NAN", "", "全部"]:
        return matches['學分數'].sum() >= float(req_row['學分數'])
    allowed_list = [d.strip() for d in allowed_val.split(',')]
    if '開課單位代碼' in matches.columns:
        valid_matches = matches[matches['開課單位代碼'].str.upper().isin(allowed_list)]
        return valid_matches['學分數'].sum() >= float(req_row['學分數'])
    return False

# --- 4. 主程式 ---
DATA_FILE = "master_data.xlsx"
if os.path.exists(DATA_FILE):
    df_courses, df_summary, prog_map = load_all_data(DATA_FILE)
else:
    st.error("找不到 master_data.xlsx"); st.stop()

tab_browse, tab_search, tab_audit = st.tabs(["📂 按學程瀏覽", "🔍 依課程反查學程", "🎓 學程完成度自動比對"])

# --- TAB 1: 瀏覽 (保持不變) ---
with tab_browse:
    st.header("學分學程規範查詢")
    if "b1" not in st.session_state: st.session_state.b1 = "全部"
    c_b1, c_b2, c_b3, c_res = st.columns([2, 2, 2, 1])
    with c_b1:
        sel_col = st.selectbox("1. 選擇學院", ["全部"] + sorted(df_courses["學院"].dropna().unique().tolist()), key="b1")
    with c_b2:
        temp_p = df_courses if sel_col == "全部" else df_courses[df_courses["學院"] == sel_col]
        sel_prog = st.selectbox("2. 選擇學程", sorted(temp_p["學程名稱"].dropna().unique().tolist()))
    with c_b3:
        year_list = sorted(df_courses[df_courses["學程名稱"]==sel_prog]["適用年度"].unique(), reverse=True)
        sel_year = st.selectbox("3. 選擇年度", year_list)
    with c_res:
        st.markdown("<br>", unsafe_allow_html=True)
        st.button("🔄 重設篩選", on_click=reset_filters, use_container_width=True)

    prog_data = df_courses[(df_courses["學程名稱"] == sel_prog) & (df_courses["適用年度"] == sel_year)]
    sum_row = df_summary[(df_summary["學程名稱"] == sel_prog) & (df_summary["適用年度"] == sel_year)]
    if not sum_row.empty:
        s = sum_row.iloc[0]
        st.success(f"**畢業門檻：** 必修 {s.get('必修總學分',0)} / 選修 {s.get('選修總學分',0)} / 總計 {s.get('總計應修學分',0)} 學分")
        note = s.get('備註 (模組要求)', '')
        if pd.notna(note) and str(note).strip():
            st.markdown(f'<div class="note-box"><b>📝 備註 (模組要求)：</b><br>{note}</div>', unsafe_allow_html=True)

    for cat in ["必修", "選修"]:
        cat_df = prog_data[prog_data["科目類別"] == cat]
        if not cat_df.empty:
            st.markdown(f"### 📍 {cat}課程清單")
            for mod in cat_df["模組名稱"].unique():
                mod_df = cat_df[cat_df["模組名稱"] == mod]
                st.markdown(f'<div class="module-card"><div class="module-title">🔹 {mod}</div>', unsafe_allow_html=True)
                cols = [c for c in ['課程代碼', '課程名稱', '學分數', '課程認抵範圍', '認抵單位代碼'] if c in mod_df.columns]
                st.dataframe(mod_df[cols], use_container_width=True, hide_index=True)
                st.markdown('</div>', unsafe_allow_html=True)

# --- TAB 2: 搜尋 (保持不變) ---
with tab_search:
    st.header("🔍 搜尋課程找學程")
    query = st.text_input("📝 請輸入課程名稱或代碼")
    if query:
        search_res = df_courses[df_courses["課程名稱"].str.contains(query, case=False, na=False) | 
                                df_courses["課程代碼"].str.contains(query, case=False, na=False)].copy()
        if not search_res.empty:
            st.dataframe(search_res[['學院', '學程名稱', '適用年度', '課程代碼', '課程名稱']].drop_duplicates(), use_container_width=True, hide_index=True)
        else: st.warning("查無資料")

# --- TAB 3: 修改重點 - 分模組顯示比對結果 ---
with tab_audit:
    st.header("🎓 學程完成度自動比對")
    
    col_dl, col_ul = st.columns([1, 2])
    with col_dl:
        st.download_button("📥 下載成績單範本", data=get_template_excel(), file_name="成績單範本.xlsx", use_container_width=True)
    with col_ul:
        uploaded_file = st.file_uploader("上傳成績單 (Excel)", type=["xlsx"])

    selected_audit_progs = st.multiselect("請勾選欲比對的學程", sorted(df_summary['學程名稱'].dropna().unique().tolist()))

    if uploaded_file:
        try:
            user_df = pd.read_excel(uploaded_file, dtype=str).dropna(how='all')
            user_df.columns = user_df.columns.str.strip()
            user_df['學分數'] = pd.to_numeric(user_df['學分數'], errors='coerce').fillna(0.0)
            user_df['及格'] = user_df['成績'].apply(check_passing)
            passed_df = user_df[user_df['及格']]

            for p_name in selected_audit_progs:
                p_sum_row = df_summary[df_summary['學程名稱'] == p_name].sort_values('適用年度', ascending=False).iloc[0]
                p_courses = df_courses[(df_courses['學程名稱'] == p_name) & (df_courses['適用年度'] == p_sum_row['適用年度'])].copy()
                
                # 執行比對
                p_courses['已完成'] = p_courses.apply(lambda r: check_course_completion(r, passed_df), axis=1)
                
                # 計算學分
                done_total = p_courses[p_courses['已完成']]['學分數'].sum()
                goal_total = float(p_sum_row['總計應修學分'])
                pct = min(done_total / goal_total, 1.0) if goal_total > 0 else 0.0

                # 顯示學程卡片
                st.markdown('<div class="prog-container">', unsafe_allow_html=True)
                ca, cb = st.columns([4, 1])
                ca.markdown(f"### 🎓 {p_name} ({p_sum_row['適用年度']}年度)")
                ca.progress(pct)
                cb.markdown(f"## {int(pct*100)}%")
                
                # 顯示備註
                p_note = p_sum_row.get('備註 (模組要求)', '')
                if pd.notna(p_note) and str(p_note).strip():
                    st.info(f"📌 **模組要求說明：** {p_note}")
                
                st.write(f"📊 已達成總學分：{int(done_total)} / 門檻：{int(goal_total)}")
                
                # 【修改核心：分模組顯示明細】
                with st.expander("🔍 查看各模組認抵明細"):
                    # 依模組分組
                    for mod_name in p_courses['模組名稱'].unique():
                        mod_data = p_courses[p_courses['模組名稱'] == mod_name].copy()
                        mod_done_credits = mod_data[mod_data['已完成']]['學分數'].sum()
                        
                        # 模組標題列
                        st.markdown(f'<div class="audit-module-header">🔹 模組：{mod_name} (已獲得 {int(mod_done_credits)} 學分)</div>', unsafe_allow_html=True)
                        
                        # 整理並顯示該模組的課程清單
                        display_df = mod_data[['科目類別', '課程代碼', '課程名稱', '學分數', '已完成']].copy()
                        display_df['已完成'] = display_df['已完成'].map({True: "✅ 已達成", False: "❌ 未完成"})
                        
                        st.table(display_df)
                
                st.markdown('</div>', unsafe_allow_html=True)
        except Exception as e:
            st.error(f"分析失敗：{e}")