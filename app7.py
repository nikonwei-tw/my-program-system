import streamlit as st
import pandas as pd
import os
import io

st.set_page_config(page_title="學分學程查詢與比對系統", layout="wide")

# --- 1. 自定義 CSS (強化進度條與卡片樣式) ---
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
        box-shadow: 2px 2px 8px rgba(0,0,0,0.08);
    }
    .module-title {
        color: #0056b3;
        font-weight: bold;
        font-size: 1.15rem;
    }
    </style>
    """, unsafe_allow_html=True)

# --- 2. 資料讀取與核心邏輯 ---
@st.cache_data
def load_all_data(file_path):
    try:
        all_sheets = pd.read_excel(file_path, sheet_name=None, dtype=str)
        df_courses = all_sheets["科目表"]
        # 強制轉學分為數字，防錯處理
        df_courses['學分數'] = pd.to_numeric(df_courses['學分數'], errors='coerce').fillna(0)
        df_courses['模組名稱'] = df_courses.get('模組名稱', "一般科目").fillna("一般科目")
        
        df_summary = all_sheets.get("學程規範總額", pd.DataFrame())
        # 清洗總額表學分欄位
        for col in ['必修總學分', '選修總學分', '總計應修學分']:
            if col in df_summary.columns:
                df_summary[col] = pd.to_numeric(df_summary[col], errors='coerce').fillna(0.0)
        return df_courses, df_summary
    except Exception as e:
        st.error(f"讀取 Excel 失敗：{e}"); st.stop()

def check_passing(grade):
    """判斷成績是否合格 (支援分數與文字)"""
    grade_str = str(grade).strip().upper()
    if grade_str in ["通過", "及格", "P", "PASS"]: return True
    try:
        return float(grade_str) >= 60
    except:
        return False

def get_template():
    """產生提供給使用者填寫的成績單範本"""
    df_temp = pd.DataFrame(columns=['課程代碼', '課程名稱', '學分數', '開課單位', '成績'])
    df_temp.loc[0] = ['CS101', '範例課程A', '3', '資工系', '85']
    df_temp.loc[1] = ['GE202', '範例課程B', '2', '通識中心', '通過']
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_temp.to_excel(writer, index=False)
    return output.getvalue()

# --- 3. 介面顯示函數 ---
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

# --- 4. 主程式執行 ---
DATA_FILE = "master_data.xlsx"
if os.path.exists(DATA_FILE):
    df_courses, df_summary = load_all_data(DATA_FILE)
else:
    st.error("找不到 master_data.xlsx"); st.stop()

tab_browse, tab_search, tab_audit = st.tabs(["📂 按學程瀏覽", "🔍 課程反查", "🎓 畢業達成度比對"])

# --- TAB 1 & 2 (保持原有功能) ---
with tab_browse:
    col_b1, col_b2, col_b3, _ = st.columns([2, 2, 2, 1])
    with col_b1: sel_col = st.selectbox("1. 選擇學院", ["全部"] + sorted(df_courses["學院"].dropna().unique().tolist()))
    with col_b2:
        temp_p = df_courses if sel_col == "全部" else df_courses[df_courses["學院"] == sel_col]
        sel_prog = st.selectbox("2. 選擇學程", sorted(temp_p["學程名稱"].unique()))
    with col_b3:
        year_list = sorted(df_courses[df_courses["學程名稱"]==sel_prog]["適用年度"].dropna().unique(), reverse=True)
        sel_year = st.selectbox("3. 選擇年度", year_list)
    
    prog_data = df_courses[(df_courses["學程名稱"] == sel_prog) & (df_courses["適用年度"] == sel_year)]
    display_styled_tables(prog_data)

with tab_search:
    query = st.text_input("📝 請輸入課程名稱或代碼關鍵字")
    if query:
        res = df_courses[df_courses["課程名稱"].str.contains(query, na=False) | df_courses["課程代碼"].str.contains(query, na=False)]
        st.dataframe(res[['學院', '學程名稱', '適用年度', '課程代碼', '課程名稱']].drop_duplicates(), use_container_width=True, hide_index=True)

# --- 5. TAB 3: 畢業達成度比對 (新增功能) ---
with tab_audit:
    st.header("🎓 多學程完成度自動審查")
    
    # 範本下載
    st.subheader("第一步：下載並填寫成績單")
    st.download_button("📥 下載成績單範本", data=get_template(), file_name="成績單範本.xlsx")
    
    # 上傳與學程選擇
    st.subheader("第二步：上傳並選擇比對對象")
    uploaded_file = st.file_uploader("上傳您的成績單 (Excel)", type=["xlsx"])
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
            st.info(f"✅ 成功讀取成績單，包含 {len(passed_df)} 門及格課程。")

            for p_name in selected_progs:
                # 抓取該學程最新年度規範
                p_sum_all = df_summary[df_summary['學程名稱'] == p_name].sort_values('適用年度', ascending=False)
                if p_sum_all.empty: continue
                
                p_sum = p_sum_all.iloc[0]
                p_courses = df_courses[(df_courses['學程名稱'] == p_name) & (df_courses['適用年度'] == p_sum['適用年度'])].copy()
                
                # 比對課程是否完成
                def check_done(row):
                    match = passed_df[passed_df['課程代碼'].str.strip() == str(row['課程代碼']).strip()]
                    if not match.empty:
                        return match['學分數'].sum() >= float(row['學分數'])
                    return False

                p_courses['已完成'] = p_courses.apply(check_done, axis=1)
                
                # 學分計算
                d_req = float(p_courses[p_courses['已完成'] & (p_courses['科目類別'] == '必修')]['學分數'].sum())
                d_elec = float(p_courses[p_courses['已完成'] & (p_courses['科目類別'] == '選修')]['學分數'].sum())
                
                g_req, g_elec, g_total = float(p_sum['必修總學分']), float(p_sum['選修總學分']), float(p_sum['總計應修學分'])
                cur_total = d_req + d_elec
                pct = min(cur_total / g_total, 1.0) if g_total > 0 else 0.0

                # 顯示單個學程進度卡片
                st.markdown('<div class="prog-container">', unsafe_allow_html=True)
                ca, cb = st.columns([4, 1])
                ca.markdown(f"### 🎓 {p_name} ({p_sum['適用年度']})")
                ca.progress(pct)
                cb.markdown(f"## {int(pct*100)}%")
                
                m1, m2, m3 = st.columns(3)
                m1.metric("必修進度", f"{int(d_req)}/{int(g_req)}")
                m2.metric("選修進度", f"{int(d_elec)}/{int(g_elec)}")
                m3.metric("目前總計", f"{int(cur_total)}/{int(g_total)}")
                
                # 缺額提示
                s_req, s_elec = max(0, g_req - d_req), max(0, g_elec - d_elec)
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
            st.error(f"分析失敗，請檢查檔案格式。錯誤：{e}")