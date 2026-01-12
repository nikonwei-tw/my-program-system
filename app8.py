import streamlit as st
import pandas as pd
import os
import io
import re

# --- 0. 頁面設定 ---
st.set_page_config(page_title="學分學程查詢與比對系統", layout="wide")

# --- 1. 自定義 CSS (維持您的視覺風格) ---
st.markdown("""
    <style>
    .module-card {
        background-color: #f8f9fa; border-radius: 10px; padding: 18px;
        margin-bottom: 20px; border-left: 6px solid #007bff;
    }
    .prog-container {
        background-color: #ffffff; border: 1px solid #e0e0e0;
        border-radius: 10px; padding: 20px; margin-bottom: 15px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
    }
    .module-summary-box {
        border-radius: 8px; padding: 12px 18px; margin: 10px 0;
        display: flex; justify-content: space-between; align-items: center;
        font-weight: bold;
    }
    .status-done {
        background-color: #d4edda; color: #155724;
        border-left: 6px solid #28a745;
    }
    .status-pending {
        background-color: #f1f3f5; color: #495057;
        border-left: 6px solid #adb5bd;
    }
    .note-box {
        background-color: #fff3cd; border-left: 5px solid #ffc107;
        padding: 12px; margin: 10px 0; font-size: 0.9rem; color: #856404;
    }
    </style>
    """, unsafe_allow_html=True)

# --- 2. 資料讀取 ---
@st.cache_data
def load_all_data(file_path):
    try:
        all_sheets = pd.read_excel(file_path, sheet_name=None, dtype=str)
        df_courses = all_sheets["科目表"]
        df_courses.columns = df_courses.columns.str.strip()
        df_courses['學分數'] = pd.to_numeric(df_courses['學分數'], errors='coerce').fillna(0)
        df_courses['模組名稱'] = df_courses.get('模組名稱', "一般科目").fillna("一般科目")
        
        # 建立學程代碼對照表
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

# --- 3. 核心工具函數 ---
def parse_required_credits(mod_name):
    """提取模組名稱括號中的數字，如 '必修模組(8)' -> 8.0"""
    match = re.search(r'\((\d+\.?\d*)\)', mod_name)
    return float(match.group(1)) if match else 0.0

def reset_filters():
    st.session_state.b1 = "全部"

def check_passing(grade):
    g = str(grade).strip().upper()
    if g in ["通過", "及格", "P", "PASS"]: return True
    try: return float(g) >= 60
    except: return False

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

# --- 4. 主程式執行 ---
DATA_FILE = "master_data.xlsx"
if os.path.exists(DATA_FILE):
    df_courses, df_summary, prog_map = load_all_data(DATA_FILE)
else:
    st.error("找不到資料檔 master_data.xlsx"); st.stop()

tab_browse, tab_search, tab_audit = st.tabs(["📂 按學程瀏覽", "🔍 依課程反查學程", "🎓 學程完成度自動比對"])

# --- TAB 1 & 2 (保持原有功能，欄位不變) ---
with tab_browse:
    if "b1" not in st.session_state: st.session_state.b1 = "全部"
    c_b1, c_b2, c_b3, c_res = st.columns([2, 2, 2, 1])
    with c_b1: sel_col = st.selectbox("1. 選擇學院", ["全部"] + sorted(df_courses["學院"].dropna().unique().tolist()), key="b1")
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
        st.success(f"**門檻：** 總計 {s.get('總計應修學分',0)} 學分")
        note = s.get('備註 (模組要求)', '')
        if pd.notna(note) and str(note).strip(): st.markdown(f'<div class="note-box">{note}</div>', unsafe_allow_html=True)
    
    for cat in ["必修", "選修"]:
        cat_df = prog_data[prog_data["科目類別"] == cat]
        if not cat_df.empty:
            st.markdown(f"### 📍 {cat}課程清單")
            for mod in cat_df["模組名稱"].unique():
                mod_df = cat_df[cat_df["模組名稱"] == mod]
                st.markdown(f'<div class="module-card"><div class="module-title">🔹 {mod}</div>', unsafe_allow_html=True)
                st.dataframe(mod_df[['課程代碼', '課程名稱', '學分數']], use_container_width=True, hide_index=True)
                st.markdown('</div>', unsafe_allow_html=True)

with tab_search:
    query = st.text_input("📝 請輸入關鍵字搜尋")
    if query:
        res = df_courses[df_courses["課程名稱"].str.contains(query, case=False, na=False)]
        st.dataframe(res[['學院', '學程名稱', '適用年度', '課程代碼', '課程名稱']].drop_duplicates(), use_container_width=True, hide_index=True)

# --- TAB 3: 修改重點 (欄位完整顯示與顏色邏輯) ---
with tab_audit:
    st.header("🎓 學程達成度排行比對")
    uploaded_file = st.file_uploader("上傳您的成績單 Excel", type=["xlsx"])

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
                total_done = p_courses[p_courses['已完成']]['學分數'].sum()
                goal_total = float(p_sum_row['總計應修學分'])
                pct = min(total_done / goal_total, 1.0) if goal_total > 0 else 0.0
                
                audit_results.append({
                    "name": p_name, "year": p_year, "pct": pct, 
                    "done": total_done, "goal": goal_total,
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
                    
                    with st.expander(f"🔍 模組達成明細 (總學分：{int(res['done'])}/{int(res['goal'])})"):
                        if pd.notna(res['note']) and str(res['note']).strip():
                            st.info(f"📌 **規範：** {res['note']}")
                        
                        for mod_name in res['details']['模組名稱'].unique():
                            mod_data = res['details'][res['details']['模組名稱'] == mod_name]
                            # 修復：正確加總已完成課程的學分
                            mod_done = mod_data.loc[mod_data['已完成'], '學分數'].sum()
                            
                            req_credits = parse_required_credits(mod_name)
                            is_satisfied = (mod_done >= req_credits) if req_credits > 0 else (mod_done > 0)
                            
                            status_class = "status-done" if is_satisfied else "status-pending"
                            status_icon = "✅" if is_satisfied else "⌛"
                            
                            p_text = f"已達標 ({int(mod_done)}/{int(req_credits)})" if is_satisfied else f"未完成 ({int(mod_done)}/{int(req_credits)})"
                            if req_credits == 0: p_text = f"累計 {int(mod_done)} 學分"

                            st.markdown(f"""
                                <div class="module-summary-box {status_class}">
                                    <span>{status_icon} {mod_name}</span>
                                    <span>{p_text}</span>
                                </div>
                            """, unsafe_allow_html=True)
                            
                            # 【欄位設定】：依照要求顯示特定欄位
                            show_cols = ['科目類別', '課程代碼', '課程名稱', '學分數', '認抵單位代碼', '備註', '互斥代碼', '已完成']
                            # 過濾掉資料庫中不存在的欄位避免報錯
                            actual_cols = [c for c in show_cols if c in mod_data.columns]
                            
                            sub_df = mod_data[actual_cols].copy()
                            sub_df['狀態'] = sub_df['已完成'].map({True: "✅ 已達成", False: "❌ 未達成"})
                            
                            # 最終顯示欄位順序
                            final_display = [c for c in actual_cols if c != '已完成'] + ['狀態']
                            st.table(sub_df[final_display])
                            
                    st.markdown('</div>', unsafe_allow_html=True)
        except Exception as e:
            st.error(f"分析失敗：{e}")