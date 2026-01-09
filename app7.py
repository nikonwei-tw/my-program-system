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

# --- 2. 資料讀取與清洗 (防錯強化) ---
@st.cache_data
def load_all_data(file_path):
    try:
        all_sheets = pd.read_excel(file_path, sheet_name=None, dtype=str)
        
        # 讀取科目表
        df_courses = all_sheets.get("科目表", pd.DataFrame())
        if df_courses.empty:
            st.error("Excel 中找不到 '科目表' 分頁"); st.stop()
            
        df_courses['學分數'] = pd.to_numeric(df_courses['學分數'], errors='coerce').fillna(0)
        df_courses['模組名稱'] = df_courses.get('模組名稱', "一般科目").fillna("一般科目")
        
        # 讀取規範總額表
        df_summary = all_sheets.get("學程規範總額", pd.DataFrame())
        if df_summary.empty:
            st.error("Excel 中找不到 '學程規範總額' 分頁"); st.stop()
            
        # 清洗總額表學分欄位
        num_cols = ['必修總學分', '選修總學分', '總計應修學分']
        for col in num_cols:
            if col in df_summary.columns:
                df_summary[col] = pd.to_numeric(df_summary[col], errors='coerce').fillna(0.0)
            else:
                df_summary[col] = 0.0 # 若缺少欄位則補 0
        
        return df_courses, df_summary
    except Exception as e:
        st.error(f"讀取 Excel 失敗：{e}"); st.stop()

# 輔助功能
def check_passing(grade):
    grade_str = str(grade).strip().upper()
    if grade_str in ["通過", "及格", "P", "PASS"]: return True
    try:
        return float(grade_str) >= 60
    except:
        return False

def get_template():
    df_temp = pd.DataFrame(columns=['課程代碼', '課程名稱', '學分數', '開課單位', '成績'])
    df_temp.loc[0] = ['CS101', '程式設計', '3', '資工系', '85']
    df_temp.loc[1] = ['GE202', '溝通技巧', '2', '通識中心', '通過']
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_temp.to_excel(writer, index=False)
    return output.getvalue()

# --- 3. 主程式 ---
DATA_FILE = "master_data.xlsx"
if os.path.exists(DATA_FILE):
    df_courses, df_summary = load_all_data(DATA_FILE)
else:
    st.error(f"找不到檔案: {DATA_FILE}"); st.stop()

tab_browse, tab_search, tab_audit = st.tabs(["📂 按學程瀏覽", "🔍 課程反查", "🎓 畢業達成度比對"])

# --- TAB 1 & 2 簡化處理 ---
with tab_browse:
    col_b1, col_b2, col_b3, _ = st.columns([2, 2, 2, 1])
    with col_b1:
        college_list = ["全部"] + sorted(df_courses["學院"].dropna().unique().tolist())
        sel_col = st.selectbox("選擇學院", college_list)
    with col_b2:
        temp_p = df_courses if sel_col == "全部" else df_courses[df_courses["學院"] == sel_col]
        prog_list = sorted(temp_p["學程名稱"].dropna().unique().tolist())
        sel_prog = st.selectbox("選擇學程", prog_list)
    with col_b3:
        year_list = sorted(df_courses[df_courses["學程名稱"]==sel_prog]["適用年度"].dropna().unique(), reverse=True)
        sel_year = st.selectbox("選擇年度", year_list)
    
    prog_data = df_courses[(df_courses["學程名稱"] == sel_prog) & (df_courses["適用年度"] == sel_year)]
    st.dataframe(prog_data[['科目類別', '模組名稱', '課程代碼', '課程名稱', '學分數']], use_container_width=True, hide_index=True)

with tab_search:
    query = st.text_input("搜尋課程關鍵字")
    if query:
        res = df_courses[df_courses["課程名稱"].str.contains(query, na=False) | df_courses["課程代碼"].str.contains(query, na=False)]
        st.dataframe(res[['學院', '學程名稱', '適用年度', '課程代碼', '課程名稱']].drop_duplicates(), use_container_width=True, hide_index=True)

# --- 4. TAB 3: 畢業達成度比對 (修正 TypeError 處) ---
with tab_audit:
    st.header("🎓 學程完成度自動審查")
    
    # 下載範本按鈕
    st.download_button("📥 下載成績單範本", data=get_template(), file_name="成績單範本.xlsx")
    
    uploaded_file = st.file_uploader("上傳您的成績單 (Excel)", type=["xlsx"])
    
    # 修正重點：使用 dropna() 確保 sorted 不會出錯
    available_progs = sorted(df_summary['學程名稱'].dropna().unique().tolist())
    selected_progs = st.multiselect("勾選欲比對的學程", available_progs, default=available_progs)

    if uploaded_file:
        try:
            # 讀取學生資料
            user_df = pd.read_excel(uploaded_file, dtype=str).dropna(how='all')
            user_df['學分數'] = pd.to_numeric(user_df['學分數'], errors='coerce').fillna(0.0)
            user_df['及格'] = user_df['成績'].apply(check_passing)
            passed_df = user_df[user_df['及格'] == True]
            
            st.divider()

            for p_name in selected_progs:
                # 抓取該學程規範
                p_sum_all = df_summary[df_summary['學程名稱'] == p_name].sort_values('適用年度', ascending=False)
                if p_sum_all.empty: continue
                
                p_sum = p_sum_all.iloc[0]
                # 抓取該學程最新年度科目表
                p_courses = df_courses[(df_courses['學程名稱'] == p_name) & (df_courses['適用年度'] == p_sum['適用年度'])].copy()
                
                # 比對函數
                def check_done(row):
                    match = passed_df[passed_df['課程代碼'].str.strip() == str(row['課程代碼']).strip()]
                    if not match.empty:
                        return match['學分數'].sum() >= float(row['學分數'])
                    return False

                p_courses['已完成'] = p_courses.apply(check_done, axis=1)
                
                # 學分加總 (確保轉為 float)
                d_req = float(p_courses[p_courses['已完成'] & (p_courses['科目類別'] == '必修')]['學分數'].sum())
                d_elec = float(p_courses[p_courses['已完成'] & (p_courses['科目類別'] == '選修')]['學分數'].sum())
                
                g_req = float(p_sum['必修總學分'])
                g_elec = float(p_sum['選修總學分'])
                g_total = float(p_sum['總計應修學分'])
                
                cur_total = d_req + d_elec
                pct = min(cur_total / g_total, 1.0) if g_total > 0 else 0.0

                # UI 顯示
                st.markdown(f'<div class="prog-container">', unsafe_allow_html=True)
                ca, cb = st.columns([4, 1])
                ca.markdown(f"### 🎓 {p_name} ({p_sum['適用年度']})")
                ca.progress(pct)
                cb.markdown(f"## {int(pct*100)}%")
                
                m1, m2, m3 = st.columns(3)
                m1.metric("必修進度", f"{int(d_req)}/{int(g_req)}")
                m2.metric("選修進度", f"{int(d_elec)}/{int(g_elec)}")
                m3.metric("目前總計", f"{int(cur_total)}/{int(g_total)}")
                
                # 缺額提示
                s_req = max(0, g_req - d_req)
                s_elec = max(0, g_elec - d_elec)
                if cur_total >= g_total and s_req == 0:
                    st.success("🎉 已達成畢業門檻！")
                else:
                    st.warning(f"⚠️ 尚缺：必修 {int(s_req)} 學分 / 選修 {int(s_elec)} 學分")
                
                with st.expander("查看比對明細"):
                    detail = p_courses[['已完成', '科目類別', '課程代碼', '課程名稱', '學分數']].copy()
                    detail['已完成'] = detail['已完成'].map({True: "✅", False: "❌"})
                    st.table(detail)
                st.markdown('</div>', unsafe_allow_html=True)

        except Exception as e:
            st.error(f"分析失敗，請檢查 Excel 格式是否正確。錯誤訊息：{e}")