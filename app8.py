import streamlit as st
import pandas as pd
import os
import io
import re  # 導入正規表達式用於解析括號

# --- 0. 頁面設定 ---
st.set_page_config(page_title="學分學程查詢與比對系統", layout="wide")

# --- 1. 自定義 CSS ---
st.markdown("""
    <style>
    .prog-container {
        background-color: #ffffff; border: 1px solid #e0e0e0;
        border-radius: 10px; padding: 20px; margin-bottom: 15px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
    }
    .module-summary-box {
        border-radius: 8px; padding: 10px 15px; margin: 8px 0;
        display: flex; justify-content: space-between; align-items: center;
        font-weight: bold;
    }
    .status-done {
        background-color: #d4edda; color: #155724;
        border-left: 6px solid #28a745;
    }
    .status-pending {
        background-color: #f1f3f5; color: #6c757d;
        border-left: 6px solid #adb5bd;
    }
    .note-box {
        background-color: #fff3cd; border-left: 5px solid #ffc107;
        padding: 12px; margin: 10px 0; font-size: 0.9rem; color: #856404;
    }
    </style>
    """, unsafe_allow_html=True)

# --- 2. 資料讀取與處理 ---
@st.cache_data
def load_all_data(file_path):
    try:
        all_sheets = pd.read_excel(file_path, sheet_name=None, dtype=str)
        df_courses = all_sheets["科目表"]
        df_courses.columns = df_courses.columns.str.strip()
        df_courses['學分數'] = pd.to_numeric(df_courses['學分數'], errors='coerce').fillna(0)
        df_courses['模組名稱'] = df_courses.get('模組名稱', "一般科目").fillna("一般科目")
        
        prog_map = df_courses[['學程代碼', '學程名稱']].drop_duplicates().set_index('學程代碼')['學程名稱'].to_dict()
        
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

# --- 3. 輔助功能 ---
def check_passing(grade):
    g = str(grade).strip().upper()
    if g in ["通過", "及格", "P", "PASS"]: return True
    try: return float(g) >= 60
    except: return False

def extract_req_credits(mod_name):
    """解析模組名稱中的括號學分，例如 '選修模組(10)' -> 10.0"""
    match = re.search(r'\((\d+)\)', str(mod_name))
    if match:
        return float(match.group(1))
    return 0.0

def check_course_completion(req_row, passed_df):
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
    st.error("找不到資料檔"); st.stop()

tab_browse, tab_search, tab_audit = st.tabs(["📂 按學程瀏覽", "🔍 依課程反查學程", "🎓 學程完成度自動比對"])

# (Tab 1 & 2 保持不變，略過以縮短篇幅，實際執行時請保留前幾次版本內容)
with tab_browse:
    st.info("請於此分頁瀏覽學程規範...")

with tab_search:
    st.info("請於此分頁搜尋課程...")

# --- TAB 3: 重點優化判定邏輯 ---
with tab_audit:
    st.header("🎓 學程達成度全局排行")
    uploaded_file = st.file_uploader("請上傳您的成績單 (Excel)", type=["xlsx"])

    if uploaded_file:
        try:
            user_df = pd.read_excel(uploaded_file, dtype=str).dropna(how='all')
            user_df.columns = user_df.columns.str.strip()
            user_df['學分數'] = pd.to_numeric(user_df['學分數'], errors='coerce').fillna(0.0)
            user_df['及格'] = user_df['成績'].apply(check_passing)
            passed_df = user_df[user_df['及格']]

            audit_results = []
            for _, p_sum_row in df_summary.iterrows():
                p_name = p_sum_row['學程名稱']
                p_year = p_sum_row['適用年度']
                p_courses = df_courses[(df_courses['學程名稱'] == p_name) & (df_courses['適用年度'] == p_year)].copy()
                if p_courses.empty: continue
                
                p_courses['已完成'] = p_courses.apply(lambda r: check_course_completion(r, passed_df), axis=1)
                done_total = p_courses[p_courses['已完成']]['學分數'].sum()
                goal_total = float(p_sum_row['總計應修學分'])
                pct = min(done_total / goal_total, 1.0) if goal_total > 0 else 0.0
                
                audit_results.append({
                    "name": p_name, "year": p_year, "pct": pct, 
                    "done": done_total, "goal": goal_total,
                    "note": p_sum_row.get('備註 (模組要求)', ''),
                    "details": p_courses
                })

            audit_results.sort(key=lambda x: x['pct'], reverse=True)

            for res in audit_results:
                with st.container():
                    st.markdown('<div class="prog-container">', unsafe_allow_html=True)
                    col_t, col_p = st.columns([4, 1])
                    col_t.markdown(f"#### 🎓 {res['name']} ({res['year']}年度)")
                    col_p.markdown(f"### {int(res['pct']*100)}%")
                    st.progress(res['pct'])
                    
                    with st.expander(f"🔍 達成明細 (總學分：{int(res['done'])} / {int(res['goal'])})"):
                        if pd.notna(res['note']) and str(res['note']).strip():
                            st.warning(f"📌 **規則說明：** {res['note']}")
                        
                        # --- 核心改動：按模組解析進度 ---
                        for mod_name in res['details']['模組名稱'].unique():
                            mod_data = res['details'][res['details']['模組名稱'] == mod_name]
                            mod_done = mod_data[mod_data['已完成']]['學分數'].sum()
                            
                            # 1. 提取門檻 (例如從 "選修模組(10)" 提取 10)
                            mod_required = extract_req_credits(mod_name)
                            
                            # 2. 判斷是否達標
                            # 如果有設定門檻(>0)，則需 已獲 >= 門檻 才是綠色
                            # 如果沒設定門檻(0)，則 已獲 > 0 就算綠色
                            if mod_required > 0:
                                is_completed = mod_done >= mod_required
                                display_text = f"{int(mod_done)} / {int(mod_required)} 學分"
                            else:
                                is_completed = mod_done > 0
                                display_text = f"已獲 {int(mod_done)} 學分"

                            status_class = "status-done" if is_completed else "status-pending"
                            status_icon = "✅" if is_completed else "⚪"
                            
                            st.markdown(f"""
                                <div class="module-summary-box {status_class}">
                                    <span>{status_icon} 模組：{mod_name}</span>
                                    <span>進度：{display_text}</span>
                                </div>
                            """, unsafe_allow_html=True)
                            
                            sub_df = mod_data[['科目類別', '課程代碼', '課程名稱', '學分數', '已完成']].copy()
                            sub_df['狀態'] = sub_df['已完成'].map({True: "✅ 已達成", False: "❌ 未達成"})
                            st.table(sub_df[['科目類別', '課程代碼', '課程名稱', '學分數', '狀態']])
                            
                    st.markdown('</div>', unsafe_allow_html=True)
        except Exception as e:
            st.error(f"分析失敗：{e}")