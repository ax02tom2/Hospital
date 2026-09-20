import io
import os
import json
import pandas as pd
import streamlit as st

# 設定網頁標題與寬度
st.set_page_config(
    page_title="醫療排程報表處理小幫手", page_icon="📋", layout="wide"
)

st.title("📋 醫療排程報表自動化處理與檢核系統")
st.write(
    "請上傳您的排程報表 Excel 檔案（支援 .xls / .xlsx），系統將自動過濾欄位、依照筆畫排序執行醫師、檢查門診/床號與心臟科醫師，並提供數量統計。"
)

# --- 醫師名單記憶功能 ---
DOCTORS_FILE = "doctors_list.json"
DEFAULT_DOCTORS = ["謝尚勳", "鄭品容", "王士鴻", "尹玉聰", "蔡榮庭", "柯子翔", "鍾禎智", "劉如濟"]

def load_doctors():
    """從檔案讀取醫師名單，若檔案不存在則回傳預設名單"""
    if os.path.exists(DOCTORS_FILE):
        try:
            with open(DOCTORS_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return DEFAULT_DOCTORS

def save_doctors(doctors_list):
    """將醫師名單存入檔案"""
    with open(DOCTORS_FILE, "w", encoding="utf-8") as f:
        json.dump(doctors_list, f, ensure_ascii=False)

# 初始化 session state
if "cardio_doctors" not in st.session_state:
    st.session_state.cardio_doctors = load_doctors()

# 側邊欄設定心臟科醫師名單
st.sidebar.header("⚙️ 檢核設定")
st.sidebar.subheader("❤️ 心臟科醫師名單")
st.sidebar.write("請直接在下方表格編輯、刪除或捲動到底部新增醫師：")

# 建立供表格編輯用的 DataFrame
df_doctors = pd.DataFrame({"醫師姓名": st.session_state.cardio_doctors})

# 使用互動式表格 (num_rows="dynamic" 允許使用者新增或刪除列)
edited_df = st.sidebar.data_editor(
    df_doctors,
    num_rows="dynamic",
    use_container_width=True,
    hide_index=True
)

if st.sidebar.button("💾 儲存醫師名單"):
    # 從編輯後的表格提取名單，去除空白與重複項目
    new_list = edited_df["醫師姓名"].dropna().astype(str).str.strip().tolist()
    new_list = [name for name in new_list if name] # 排除空字串
    
    seen = set()
    new_list_unique = [x for x in new_list if not (x in seen or seen.add(x))]
    
    # 更新系統狀態並存檔
    st.session_state.cardio_doctors = new_list_unique
    save_doctors(new_list_unique)
    st.sidebar.success("✅ 名單已成功儲存！下次開啟將自動載入。")

# 正式提供比對的醫師名單
cardio_doctors = st.session_state.cardio_doctors

st.sidebar.markdown("---")

# 檔案上傳區塊
uploaded_file = st.file_uploader(
    "請選擇或拖曳報表檔案 (Excel)", type=["xls", "xlsx"]
)

if uploaded_file is not None:
  try:
    # 讀取 Excel 檔案：加上 dtype=str 強制以純文字讀取，保留病歷號前面所有的 0
    if uploaded_file.name.endswith(".xls"):
      df = pd.read_excel(uploaded_file, engine="xlrd", dtype=str)
    else:
      df = pd.read_excel(uploaded_file, dtype=str)

    # 將空值補為空字串，避免出現 nan
    df = df.fillna("")

    st.success(
        f"✅ 成功載入檔案：**{uploaded_file.name}** (共 {len(df)} 筆資料)"
    )

    # 1. 欄位保留與對應
    target_columns_map = {
        "診別": "診別",
        "病例號": "病歷號",
        "姓名": "姓名",
        "性別": "性別",
        "執行日期": (
            "執行起始日期"
            if "執行起始日期" in df.columns
            else df.columns[0]
        ),
        "開單日期": "開單日" if "開單日" in df.columns else df.columns[0],
        "開單醫生": "開單醫師" if "開單醫師" in df.columns else df.columns[0],
        "執行醫生": (
            "操作醫師"
            if "操作醫師" in df.columns
            else ("預設執行醫師" if "預設執行醫師" in df.columns else df.columns[0])
        ),
    }

    processed_df = pd.DataFrame()
    for col_name, raw_col in target_columns_map.items():
      if raw_col in df.columns:
        processed_df[col_name] = df[raw_col].astype(str).str.strip()
      else:
        processed_df[col_name] = "未提供"

    # 2. 執行醫生排序
    processed_df = processed_df.sort_values(by="執行醫生").reset_index(drop=True)

    # 3. 檢查開單醫生是否為心臟科醫師
    processed_df["是否心臟科"] = processed_df["開單醫生"].isin(cardio_doctors)
    
    # 4. 處理「執行日期」去掉年份
    def remove_year(date_str):
        date_str = str(date_str).strip()
        if " " in date_str:
            date_str = date_str.split(" ")[0]
        if len(date_str) >= 8 and (date_str[4] == "/" or date_str[4] == "-"):
            return date_str[5:]
        return date_str
        
    processed_df["執行日期"] = processed_df["執行日期"].apply(remove_year)

    # ==========================================
    # 5. 網頁上方：明確顯示各個執行醫生的數量 (自訂放大樣式)
    # ==========================================
    st.markdown("### 📈 執行醫師工作量統計")
    
    doctor_counts = processed_df["執行醫生"].value_counts()
    
    # 依醫師數量動態建立欄位
    metric_cols = st.columns(len(doctor_counts))
    for i, (doc, count) in enumerate(doctor_counts.items()):
        with metric_cols[i]:
            # 使用自訂 HTML 讓醫師名字與次數一樣大、一樣醒目
            st.markdown(
                f"""
                <div style="background-color: #f0f2f6; padding: 20px; border-radius: 10px; text-align: center; box-shadow: 1px 1px 5px rgba(0,0,0,0.1);">
                    <h2 style="margin: 0; color: #31333F;">👨‍⚕️ {doc}</h2>
                    <h2 style="margin: 10px 0 0 0; color: #0068c9;">{count} 人次</h2>
                </div>
                """, 
                unsafe_allow_html=True
            )
            
    st.markdown("---")

    # ==========================================
    # 6. 網頁顯示：主表格預覽
    # ==========================================
    st.markdown("### 📊 處理後的排程報表預覽")
    st.markdown(
        "> 💡 **說明**：\n"
        "> - **病歷號** 已完整保留包含 `0` 開頭的所有數字。\n"
        "> - 診別若包含**「門」**字則顯示黑色；否則以 **紅色字體** 顯示。\n"
        "> - **執行日期** 已隱藏年份，並統一以 **紅色字體** 顯示。\n"
        "> - 若開單醫生**不在左側儲存的心臟科醫師名單內**，姓名會以 **紅色字體** 顯示並加上警示。"
    )

    # 自訂 HTML 表格渲染 (去除縮排避免變成 Markdown 程式碼區塊)
    def render_custom_table(df_data):
        html = '<div style="overflow-x: auto;">'
        html += '<style>'
        html += '.report-table { width: 100%; border-collapse: collapse; font-family: sans-serif; font-size: 14px; }'
        html += '.report-table th, .report-table td { border: 1px solid #ddd; padding: 8px 12px; text-align: center; }'
        html += '.report-table th { background-color: #f0f2f6; color: #31333F; }'
        html += '.red-text { color: red; font-weight: bold; }'
        html += '.warning-badge { background-color: #ff4b4b; color: white; padding: 2px 6px; border-radius: 4px; font-size: 12px; font-weight: bold; margin-left: 5px; }'
        html += '</style>'
        html += '<table class="report-table"><thead><tr>'
        html += '<th>項次</th><th>診別</th><th>病歷號</th><th>姓名</th><th>性別</th><th>執行日期</th><th>開單日期</th><th>開單醫生</th><th>執行醫生</th>'
        html += '</tr></thead><tbody>'

        for idx, row in df_data.iterrows():
            # --- 診別邏輯 ---
            zbie_val = str(row['診別'])
            zbie_class = "" if "門" in zbie_val else "red-text"

            # --- 開單醫生邏輯：非心臟科變紅字 ---
            doc_display = row["開單醫生"]
            doc_class = ""
            if not row["是否心臟科"]:
                doc_class = "red-text"
                doc_display = f"{doc_display} <span class='warning-badge'>⚠️ 非心臟科</span>"

            # 採用無換行的方式寫入，徹底避免 Streamlit 的程式碼方塊渲染問題
            html += f'<tr><td>{idx + 1}</td><td class="{zbie_class}">{zbie_val}</td><td>{row["病例號"]}</td><td>{row["姓名"]}</td><td>{row["性別"]}</td><td class="red-text">{row["執行日期"]}</td><td>{row["開單日期"]}</td><td class="{doc_class}">{doc_display}</td><td><b>{row["執行醫生"]}</b></td></tr>'

        html += '</tbody></table></div>'
        return html

    st.markdown(render_custom_table(processed_df), unsafe_allow_html=True)


    # ==========================================
    # 7. 匯出下載功能 (純淨報表，不含統計數量)
    # ==========================================
    st.markdown("---")
    st.markdown("### 💾 匯出處理後報表")

    export_df = processed_df[
        [
            "診別",
            "病例號",
            "姓名",
            "性別",
            "執行日期",
            "開單日期",
            "開單醫生",
            "執行醫生",
        ]
    ].copy()

    def to_excel(df_to_save):
      output = io.BytesIO()
      with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df_to_save.to_excel(writer, index=False, sheet_name="處理後報表")
        
        workbook = writer.book
        worksheet = writer.sheets["處理後報表"]
        
        # 定義紅色字體樣式
        red_format = workbook.add_format({'font_color': 'red'})
        
        # 取得需要判斷欄位的位置
        zbie_col_idx = df_to_save.columns.get_loc("診別")
        date_col_idx = df_to_save.columns.get_loc("執行日期")
        doc_col_idx = df_to_save.columns.get_loc("開單醫生")
        
        # 將紅色樣式寫入 Excel
        for row_idx in range(len(df_to_save)):
            excel_row = row_idx + 1 
            
            # 1. 診別紅字判斷 (沒有「門」字就套用紅色)
            zbie_val = str(df_to_save.iloc[row_idx, zbie_col_idx])
            if "門" not in zbie_val:
                worksheet.write_string(excel_row, zbie_col_idx, zbie_val, red_format)
                
            # 2. 執行日期全面紅字
            date_val = str(df_to_save.iloc[row_idx, date_col_idx])
            worksheet.write_string(excel_row, date_col_idx, date_val, red_format)
            
            # 3. 開單醫生紅字判斷 (不在心臟科名單內就套用紅色)
            doc_val = str(df_to_save.iloc[row_idx, doc_col_idx])
            if doc_val not in cardio_doctors:
                worksheet.write_string(excel_row, doc_col_idx, doc_val, red_format)
            
      return output.getvalue()

    excel_data = to_excel(export_df)

    st.download_button(
        label="📥 下載處理完成的 Excel 報表",
        data=excel_data,
        file_name="processed_schedule_report.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml"
            ".spreadsheetml.sheet"
        ),
    )

  except Exception as e:
    st.error(f"❌ 讀取或處理檔案時發生錯誤：{e}")
    st.info("請確認上傳的是正確格式的醫療排程報表 Excel 檔案。")
else:
  st.info("💡 請由上方按鈕上傳您的排程報表檔案 (.xls / .xlsx) 來開始使用。")
