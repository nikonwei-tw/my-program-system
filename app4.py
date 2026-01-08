import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="學分學程導航系統", layout="wide")

@st.cache_data
def load_all_data(file_path):
    try:
        all_sheets = pd.read_excel(file_path, sheet_name=None, dtype=str)
        df_courses = all_sheets["科目表"]
        df_courses['學分數'] = pd.to_numeric(df_courses['學分數'], errors='coerce').fillna(0)
        df_summary = all_sheets.get("學程規範總額", pd.DataFrame())
        return df_courses, df_summary
    except Exception as e:
        st.error(f"讀取失敗：{e}")
        st.stop()

DATA_FILE = "master_data.xlsx"
if os.path.exists(DATA_FILE):
    df_courses, df_summary = load_all_data(DATA_FILE)
else:
    st.error("找不到 master_data.xlsx")
    st.stop()

tab_browse, tab_smart_search = st.tabs(["📂 按學程瀏覽", "🎯 多課組合找學程"])

# --- TAB 1: 原有的瀏覽功能 ---
with tab_browse:
    st.header("學分學程規範查詢")
    # ... (此處保留之前的篩選與顯示邏輯，維持原樣即可) ...
    # 為了簡潔，這裡省略重複代碼，請延用上一版的 Tab 1 內容

# --- TAB 2: 多課程組合搜尋 (新功能) ---
with tab_smart_search:
    st.header("🎯 尋找我的命定學程")
    st.markdown("請輸入您**已經修過**或**有興趣**的課程名稱（多門課請用 `逗號` 分隔）：")
    
    # 輸入範例：微積分, 程式設計, 統計學
    user_input = st.text_area("課程名稱輸入區", placeholder="例如：微積分, 程式設計, 數據分析", help="輸入越多，推薦越精準！")
    
    if user_input:
        # 1. 處理輸入的關鍵字
        keywords = [k.strip() for k in user_input.replace("，", ",").split(",") if k.strip()]
        
        if keywords:
            # 2. 在全域資料中搜尋匹配的行
            # 建立一個正則表達式，例如 "微積分|程式設計|統計學"
            pattern = "|".join(keywords)
            matched_courses = df_courses[
                df_courses["課程名稱"].str.contains(pattern, case=False, na=False)
            ].copy()
            
            if not matched_courses.empty:
                # 3. 計算每個學程的命中門數
                # 以「學程名稱」與「適用年度」分組統計
                recommend_df = matched_courses.groupby(['學院', '學程名稱', '適用年度']).agg({
                    '課程名稱': lambda x: ", ".join(x.unique()), # 列出命中的課名
                    '學分數': 'sum',                             # 累計命中的學分
                    '課程代碼': 'count'                          # 計算命中的門數
                }).reset_index()
                
                recommend_df = recommend_df.rename(columns={'課程代碼': '命中門數', '課程名稱': '包含的相關課程'})
                
                # 4. 排序：命中門數越多、學分越多的排前面
                recommend_df = recommend_df.sort_values(by=['命中門數', '學分數'], ascending=False)
                
                st.success(f"根據您的課程組合，我們為您推薦以下學程：")
                
                # 5. 顯示推薦排行榜
                for index, row in recommend_df.iterrows():
                    with st.expander(f"🏆 推薦第 {index+1} 名：{row['學程名稱']} ({row['適用年度']}年度) - 命中 {row['命中門數']} 門課"):
                        c1, c2 = st.columns([1, 2])
                        with c1:
                            st.metric("累計可抵免學分", f"{int(row['學分數'])} 學分")
                        with c2:
                            st.write(f"**已包含課程：** {row['包含的相關課程']}")
                            
                        # 加一個按鈕連動回 Tab 1 (提示性質)
                        st.caption(f"提示：您可以切換到『按學程瀏覽』分頁，詳細查看『{row['學程名稱']}』的完整規範。")
                
                st.divider()
                st.write("### 🔍 詳細比對清單")
                st.dataframe(recommend_df, use_container_width=True, hide_index=True)
            else:
                st.warning("查無匹配的學程，請嘗試縮短關鍵字（例如用『微積』代替『微積分一』）。")
    else:
        st.info("💡 **操作小技巧：**\n嘗試輸入您大一、大二修過的必修課，看看是否已經不知不覺快拿到某個學程的證書了！")