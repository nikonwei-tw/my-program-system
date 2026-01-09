import streamlit as st
import pandas as pd
import os
import io

st.set_page_config(page_title="學分學程比對系統", layout="wide")

# --- 1. 自定義 CSS (加入進度條顏色優化) ---
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
    .status-pass { color: #28a745; font-weight: bold; }
    .status-fail { color: #dc3545; font-weight: bold; }
    </style>
    """, unsafe_allow_html=True)

# --- 2. 資料讀取 ---
@st.cache_data
def load_all_data(file_path):
    try:
        all_sheets = pd.read_excel(file_path, sheet_name=None, dtype=str)
        df_courses = all_sheets["科目表"]
        df_courses['學分數'] = pd.to_numeric(df_courses['學分數'], errors='coerce').fillna(0)
        df_courses['模組名稱'] = df_courses.get('模組名稱', "一般科目").fillna("一般科目")
        df_summary = all_sheets.get("學程規範總額", pd.DataFrame())
        df_summary['必修總學分'] = pd.to_numeric(df_summary['必修總學分'], errors='coerce').fillna(0)
        df_summary['選修總學分'] = pd.to_numeric(df_summary['選修總學分'], errors='coerce').fillna(0)
        df_summary['總計應修學分'] = pd.to_numeric(df_summary['總計應修學分'], errors='coerce').fillna(0)
        return df_courses, df_summary
    except Exception as e:
        st.error(f"讀取 Excel 失敗：{e}"); st.stop()

# --- 3. 功能函數：產生範本、比對邏輯 ---
def get_template():
    """產生範本檔案"""
    df_temp = pd.DataFrame(columns=['課程代碼', '課程名稱', '學分數', '開課單位', '成績'])
    # 加入一筆範例資料
    df_temp.loc[0] = ['CS101', '程式設計', '3', '資工系', '85']
    df_temp.loc[1] = ['GE202', '大數據導論', '2', '通識中心', '通過']
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_temp.to_excel(writer, index=False)
    return output.getvalue()

def check_passing(grade):
    """判斷成績是否合格"""
    grade_str = str(grade).strip()
    if grade_str == "通過":
        return True
    try:
        return float(grade_str) >= 60
    except:
        return False

# --- 4. 主程式邏輯 ---
DATA_FILE = "master_data.xlsx"
if os.path.exists(DATA_FILE):
    df_courses, df_summary = load_all_data(DATA_FILE)
else:
    st.error("找不到資料庫檔案"); st.stop()

tab_browse, tab_search, tab_audit = st.tabs(["📂 按學程瀏覽", "🔍 課程反查", "🎓 畢業達成度比對"])

# (Tab 1 & 2 保持不變，此處略過)

# --- TAB 3: 畢業達成度比對 ---
with tab_audit:
    st.header("🎓 多學程完成度自動審查")
    
    # 第一步：下載範本
    st.subheader("第一步：下載並填寫成績單")
    col_t1, col_t2 = st.columns([1, 3])
    with col_t1:
        st.download_button("📥 下載成績單範本", data=get_template(), file_name="成績單範本.xlsx")
    with col_t2:
        st.caption("※ 請依範本填寫已修課程，成績需大於 60 分或標註為「通過」始計入學分。")

    # 第二步：上傳檔案
    st.subheader("第二步：上傳您的成績單")
    uploaded_file = st.file_uploader("選擇您的成績單 Excel 檔", type=["xlsx"])
    
    # 第三步：設定比對範圍
    st.subheader("第三步：設定比對學程")
    all_progs = sorted(df_summary['學程名稱'].unique())
    selected_progs = st.multiselect("請勾選欲比對的學程（預設顯示全部）", all_progs, default=all_progs)

    if uploaded_file:
        try:
            user_df = pd.read_excel(uploaded_file, dtype=str)
            user_df['成績及格'] = user_df['成績'].apply(check_passing)
            user_df['學分數'] = pd.to_numeric(user_df['學分數'], errors='coerce').fillna(0)
            
            # 建立已及格清單 (以代碼為準)
            passed_df = user_df[user_df['成績及格'] == True]
            
            st.divider()
            st.info(f"✅ 上傳成功！偵測到 {len(passed_df)} 門合格課程。開始進行多學程比對...")

            # 遍歷所選學程
            for p_name in selected_progs:
                # 取得該學程最新年度的規範
                p_summary = df_summary[df_summary['學程名稱'] == p_name].sort_values('適用年度', ascending=False).iloc[0]
                p_courses = df_courses[(df_courses['學程名稱'] == p_name) & (df_courses['適用年度'] == p_summary['適用年度'])]
                
                # 比對邏輯：代碼需相同，且學生修的學分 >= 資料庫規定的學分
                def audit_course(row):
                    match = passed_df[passed_df['課程代碼'] == row['課程代碼']]
                    if not match.empty:
                        # 檢查學生學分是否足夠 (避免修半門課的情況)
                        if match['學分數'].astype(float).sum() >= float(row['學分數']):
                            return True
                    return False

                p_courses['是否完成'] = p_courses.apply(audit_course, axis=1)
                
                # 計算學分
                done_req = p_courses[p_courses['是否完成'] & (p_courses['科目類別'] == '必修')]['學分數'].sum()
                done_elec = p_courses[p_courses['是否完成'] & (p_courses['科目類別'] == '選修')]['學分數'].sum()
                
                goal_req = p_summary['必修總學分']
                goal_elec = p_summary['選修總學分']
                goal_total = p_summary['總計應修學分']
                current_total = done_req + done_elec
                prog_pct = min(current_total / goal_total, 1.0) if goal_total > 0 else 0

                # 顯示單個學程結果
                with st.container():
                    st.markdown(f'<div class="prog-container">', unsafe_allow_html=True)
                    c1, c2 = st.columns([3, 1])
                    with c1:
                        st.markdown(f"### 🎓 {p_name} ({p_summary['適用年度']}年度)")
                        st.progress(prog_pct)
                    with c2:
                        st.markdown(f"## {int(prog_pct*100)}%")
                    
                    # 顯示數據指標
                    m1, m2, m3 = st.columns(3)
                    m1.metric("必修進度", f"{int(done_req)}/{int(goal_req)}")
                    m2.metric("選修進度", f"{int(done_elec)}/{int(goal_elec)}")
                    
                    # 尚缺分析
                    short_req = max(0, goal_req - done_req)
                    short_elec = max(0, goal_elec - done_elec)
                    
                    msg = []
                    if short_req > 0: msg.append(f"尚缺必修 **{int(short_req)}** 學分")
                    if short_elec > 0: msg.append(f"尚缺選修 **{int(short_elec)}** 學分")
                    
                    if not msg:
                        st.success("✨ 已達成此學程所有條件！")
                    else:
                        st.warning(" ⚠️ " + "、".join(msg))
                    
                    # 展開查看細節表
                    with st.expander("查看此學程詳細比對清單"):
                        display_df = p_courses[['是否完成', '科目類別', '課程代碼', '課程名稱', '學分數', '課程認抵範圍']].copy()
                        display_df['是否完成'] = display_df['是否完成'].map({True: "✅ 已達成", False: "❌ 未達成"})
                        st.table(display_df)
                        
                    st.markdown('</div>', unsafe_allow_html=True)

        except Exception as e:
            st.error(f"分析時發生錯誤，請確認成績單格式。詳細錯誤：{e}")