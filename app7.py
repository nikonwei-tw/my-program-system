import streamlit as st
import pandas as pd
import os
import io

st.set_page_config(page_title="學分學程查詢系統", layout="wide")

# --- 1. 自定義 CSS ---
st.markdown("""
    <style>
    .prog-container {
        background-color: #ffffff;
        border: 1px solid #e0e0e0;
        border-radius: 10px;
        padding: 20px;
        margin-bottom: 25px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
    }
    .module-card {
        background-color: #f8f9fa;
        border-radius: 10px;
        padding: 18px;
        margin-bottom: 20px;
        border-left: 6px solid #007bff;
    }
    .module-title { color: #0056b3; font-weight: bold; font-size: 1.15rem; }
    </style>
    """, unsafe_allow_html=True)

# --- 2. 資料讀取與清洗 ---
@st.cache_data
def load_all_data(file_path):
    try:
        all_sheets = pd.read_excel(file_path, sheet_name=None, dtype=str)
        df_courses = all_sheets["科目表"]
        # 強制轉學分為數字，錯誤變 NaN 再補 0
        df_courses['學分數'] = pd.to_numeric(df_courses['學分數'], errors='coerce').fillna(0)
        df_courses['模組名稱'] = df_courses.get('模組名稱', "一般科目").fillna("一般科目")
        
        df_summary = all_sheets.get("學程規範總額", pd.DataFrame())
        # 清洗總額表，確保計算不報錯
        for col in ['必修總學分', '選修總學分', '總計應修學分']:
            if col in df_summary.columns:
                df_summary[col] = pd.to_numeric(df_summary[col], errors='coerce').fillna(0)
        
        return df_courses, df_summary
    except Exception as e:
        st.error(f"讀取 Excel 失敗：{e}"); st.stop()

# 輔助功能
def check_passing(grade):
    grade_str = str(grade).strip()
    if grade_str in ["通過", "及格", "P"]: return True
    try:
        return float(grade_str) >= 60
    except:
        return False

def get_template():
    df_temp = pd.DataFrame(columns=['課程代碼', '課程名稱', '學分數', '開課單位', '成績'])
    df_temp.loc[0] = ['範例CS101', '程式設計', '3', '資工系', '85']
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_temp.to_excel(writer, index=False)
    return output.getvalue()

# --- 3. 核心顯示函數 ---
def display_styled_tables(df):
    for cat in ["必修", "選修"]:
        cat_df = df[df["科目類別"] == cat]
        if not cat_df.empty:
            st.markdown(f"### 📍 {cat}課程清單")
            for mod in cat_df["模組名稱"].unique():
                mod_df = cat_df[cat_df["模組名稱"] == mod]
                st.markdown(f'<div class="module-card"><div class="module-title">🔹 {mod}</div>', unsafe_allow_html=True)
                st.dataframe(mod_df[['課程代碼', '課程名稱', '學分數', '課程認抵範圍']], use_container_width=True, hide_index=True)
                st.markdown('</div>', unsafe_allow_html=True)

# --- 4. 主程式 ---
DATA_FILE = "master_data.xlsx"
if os.path.exists(DATA_FILE):
    df_courses, df_summary = load_all_data(DATA_FILE)
else:
    st.error("找不到 master_data.xlsx"); st.stop()

tab_browse, tab_search, tab_audit = st.tabs(["📂 按學程瀏覽", "🔍 課程反查", "🎓 畢業達成度比對"])

with tab_browse:
    col_b1, col_b2, col_b3, col_reset = st.columns([2, 2, 2, 1])
    with col_b1: sel_col = st.selectbox("1. 選擇學院", ["全部"] + sorted(df_courses["學院"].dropna().unique().tolist()), key="b1")
    with col_b2: 
        temp_p = df_courses if sel_col == "全部" else df_courses[df_courses["學院"] == sel_col]
        sel_prog = st.selectbox("2. 選擇學程", sorted(temp_p["學程名稱"].unique()), key="b2")
    with col_b3: sel_year = st.selectbox("3. 選擇年度", sorted(df_courses[df_courses["學程名稱"]==sel_prog]["適用年度"].unique(), reverse=True), key="b3")
    
    prog_data = df_courses[(df_courses["學程名稱"] == sel_prog) & (df_courses["適用年度"] == sel_year)]
    summary_row = df_summary[(df_summary["學程名稱"] == sel_prog) & (df_summary["適用年度"] == sel_year)]
    
    if not summary_row.empty:
        s = summary_row.iloc[0]
        st.success(f"🎓 **畢業門檻：** 必修 {s['必修總學分']} / 選修 {s['選修總學分']} / 總計 {s['總計應修學分']} 學分")
    display_styled_tables(prog_data)

with tab_search:
    query = st.text_input("📝 請輸入課程名稱或代碼關鍵字")
    if query:
        res = df_courses[df_courses["課程名稱"].str.contains(query, na=False) | df_courses["課程代碼"].str.contains(query, na=False)].copy()
        st.dataframe(res[['學院', '學程名稱', '適用年度', '課程代碼', '課程名稱']].drop_duplicates(), use_container_width=True, hide_index=True)

# --- 核心更新：畢業進度審查 ---
with tab_audit:
    st.header("🎓 多學程完成度自動審查")
    
    # 範本下載
    st.download_button("📥 下載成績單範本", data=get_template(), file_name="成績單範本.xlsx")
    
    uploaded_file = st.file_uploader("選擇並上傳您的成績單 (Excel)", type=["xlsx"])
    all_progs = sorted(df_summary['學程名稱'].unique())
    selected_progs = st.multiselect("勾選欲比對的學程", all_progs, default=all_progs)

    if uploaded_file:
        try:
            user_df = pd.read_excel(uploaded_file, dtype=str)
            user_df['及格'] = user_df['成績'].apply(check_passing)
            user_df['學分數'] = pd.to_numeric(user_df['學分數'], errors='coerce').fillna(0)
            passed_df = user_df[user_df['及格']]
            
            st.divider()

            for p_name in selected_progs:
                # 取得該學程最新年度資料
                p_sum = df_summary[df_summary['學程名稱'] == p_name].sort_values('適用年度', ascending=False).iloc[0]
                p_courses = df_courses[(df_courses['學程名稱'] == p_name) & (df_courses['適用年度'] == p_sum['適用年度'])].copy()
                
                # 比對邏輯
                def check_done(row):
                    match = passed_df[passed_df['課程代碼'] == row['課程代碼']]
                    if not match.empty:
                        return match['學分數'].astype(float).sum() >= float(row['學分數'])
                    return False

                p_courses['已完成'] = p_courses.apply(check_done, axis=1)
                
                # 學分加總
                d_req = p_courses[p_courses['已完成'] & (p_courses['科目類別'] == '必修')]['學分數'].sum()
                d_elec = p_courses[p_courses['已完成'] & (p_courses['科目類別'] == '選修')]['學分數'].sum()
                
                g_req, g_elec, g_total = p_sum['必修總學分'], p_sum['選修總學分'], p_sum['總計應修學分']
                current_total = d_req + d_elec
                pct = min(current_total / g_total, 1.0) if g_total > 0 else 0

                # 顯示介面
                st.markdown(f'<div class="prog-container">', unsafe_allow_html=True)
                col_a, col_b = st.columns([4, 1])
                col_a.markdown(f"### 🎓 {p_name} ({p_sum['適用年度']})")
                col_a.progress(pct)
                col_b.markdown(f"## {int(pct*100)}%")
                
                m1, m2, m3 = st.columns(3)
                m1.metric("必修進度", f"{int(d_req)}/{int(g_req)}")
                m2.metric("選修進度", f"{int(d_elec)}/{int(g_elec)}")
                
                # 缺額文字
                s_req = max(0, g_req - d_req)
                s_elec = max(0, g_elec - d_elec)
                if s_req + s_elec == 0:
                    st.success("🎉 已達成畢業門檻！")
                else:
                    st.warning(f"⚠️ 尚缺：必修 {int(s_req)} 學分 / 選修 {int(s_elec)} 學分")
                
                with st.expander("查看比對明細"):
                    detail = p_courses[['已完成', '科目類別', '課程代碼', '課程名稱', '學分數']].copy()
                    detail['已完成'] = detail['已完成'].map({True: "✅", False: "❌"})
                    st.table(detail)
                st.markdown('</div>', unsafe_allow_html=True)

        except Exception as e:
            st.error(f"系統發生錯誤：{e}。請檢查 Excel 欄位是否正確。")