import io
import re
import os
import json
import base64
import pandas as pd
import streamlit as st

# 設定網頁標題與寬度
st.set_page_config(
    page_title="醫療排程報表處理小幫手", page_icon="📋", layout="wide"
)

# ==========================================
# 🖼️ 自訂桌布功能與 CSS 樣式注入
# ==========================================
st.sidebar.header("🖼️ 網頁桌布設定")
st.sidebar.write("上傳您的「塔瑪鴿子」或其他喜歡的桌布：")
bg_upload = st.sidebar.file_uploader("選擇背景圖片 (JPG/PNG)", type=["png", "jpg", "jpeg"])

if bg_upload is not None:
    # 將圖片轉為 Base64 以嵌入 CSS
    encoded_string = base64.b64encode(bg_upload.read()).decode()
    bg_ext = bg_upload.name.split('.')[-1]
    
    st.markdown(
        f"""
        <style>
        /* 替換整個網頁的背景 */
        .stApp {{
            background-image: url(data:image/{bg_ext};base64,{encoded_string});
            background-size: cover;
            background-position: center;
            background-attachment: fixed;
        }}
        /* 為了避免桌布太花讓字看不清楚，幫主要內容區塊加上半透明底色與圓角 */
        .block-container {{
            background-color: rgba(255, 255, 255, 0.90);
            padding: 2rem !important;
            border-radius: 15px;
            box-shadow: 0px 4px 15px rgba(0,0,0,0.1);
            margin-top: 2rem;
        }}
        /* 隱藏預設的頂部裝飾線條 */
        header[data-testid="stHeader"] {{
            background: transparent;
        }}
        </style>
        """,
        unsafe_allow_html=True
    )

st.title("📋 醫療排程報表自動化處理與檢核系統")
st.write(
    "請上傳您的排程報表 Excel 檔案（支援 .xls / .xlsx），系統將自動過濾欄位、依照姓氏筆畫排序執行醫師、檢查門診/床號與心臟科醫師，並提供數量統計。"
)

# --- 醫師名單記憶功能 ---
DOCTORS_FILE = "doctors_list.json"
DEFAULT_DOCTORS = ["謝尚勳", "鄭品容", "王士鴻", "尹玉聰", "蔡榮庭", "柯子翔", "鍾禎智", "劉如濟"]

def load_doctors():
    if os.path.exists(DOCTORS_FILE):
        try:
            with open(DOCTORS_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return DEFAULT_DOCTORS

def save_doctors(doctors_list):
    with open(DOCTORS_FILE, "w", encoding="utf-8") as f:
        json.dump(doctors_list, f, ensure_ascii=False)

if "cardio_doctors" not in st.session_state:
    st.session_state.cardio_doctors = load_doctors()

st.sidebar.markdown("---")
st.sidebar.header("⚙️ 檢核設定")
st.sidebar.subheader("❤️ 心臟科醫師名單")
st.sidebar.write("請直接在下方表格編輯、刪除或捲動到底部新增醫師：")

df_doctors = pd.DataFrame({"醫師姓名": st.session_state.cardio_doctors})

edited_df = st.sidebar.data_editor(
    df_doctors,
    num_rows="dynamic",
    use_container_width=True,
    hide_index=True
)

if st.sidebar.button("💾 儲存醫師名單"):
    new_list = edited_df["醫師姓名"].dropna().astype(str).str.strip().tolist()
    new_list = [name for name in new_list if name]
    
    seen = set()
    new_list_unique = [x for x in new_list if not (x in seen or seen.add(x))]
    
    st.session_state.cardio_doctors = new_list_unique
    save_doctors(new_list_unique)
    st.sidebar.success("✅ 名單已成功儲存！下次開啟將自動載入。")

cardio_doctors = st.session_state.cardio_doctors
st.sidebar.markdown("---")

# 檔案上傳區塊
uploaded_file = st.file_uploader(
    "請選擇或拖曳報表檔案 (Excel)", type=["xls", "xlsx"]
)

if uploaded_file is not None:
  try:
    if uploaded_file.name.endswith(".xls"):
      df = pd.read_excel(uploaded_file, engine="xlrd", dtype=str)
    else:
      df = pd.read_excel(uploaded_file, dtype=str)

    df = df.fillna("")
    st.success(f"✅ 成功載入檔案：**{uploaded_file.name}** (共 {len(df)} 筆資料)")

    # 1. 欄位保留與對應
    target_columns_map = {
        "診別": "診別",
        "病例號": "病歷號",
        "姓名": "姓名",
        "性別": "性別",
        "執行日期": ("執行起始日期" if "執行起始日期" in df.columns else df.columns[0]),
        "開單日期": "開單日" if "開單日" in df.columns else df.columns[0],
        "開單醫生": "開單醫師" if "開單醫師" in df.columns else df.columns[0],
        "執行醫生": ("操作醫師" if "操作醫師" in df.columns else ("預設執行醫師" if "預設執行醫師" in df.columns else df.columns[0])),
    }

    processed_df = pd.DataFrame()
    for col_name, raw_col in target_columns_map.items():
      if raw_col in df.columns:
        processed_df[col_name] = df[raw_col].astype(str).str.strip()
      else:
        processed_df[col_name] = "未提供"

    # ==========================================
    # 2. 執行醫生依照「姓氏筆畫由少到多」排序
    # ==========================================
    stroke_map = {
        '一': 1, '乙': 1,
        '丁': 2, '卜': 2, '刁': 2, '七': 2, '乃': 2,
        '于': 3, '大': 3, '山': 3, '千': 3, '上': 3, '下': 3, '凡': 3, '弓': 3,
        '王': 4, '尤': 4, '方': 4, '毛': 4, '文': 4, '仇': 4, '尹': 4, '孔': 4, '巴': 4, '水': 4, '牛': 4, '戈': 4, '支': 4, '元': 4, '公': 4, '太': 4,
        '史': 5, '石': 5, '白': 5, '田': 5, '甘': 5, '申': 5, '司': 5, '左': 5, '平': 5, '古': 5, '皮': 5, '丘': 5, '包': 5, '令': 5, '布': 5, '玉': 5, '世': 5,
        '朱': 6, '江': 6, '伍': 6, '向': 6, '安': 6, '任': 6, '吉': 6, '成': 6, '池': 6, '牟': 6, '艾': 6, '百': 6, '米': 6, '戎': 6, '伊': 6, '全': 6, '印': 6, '宇': 6, '曲': 6, '年': 6,
        '李': 7, '吳': 7, '何': 7, '呂': 7, '宋': 7, '沈': 7, '汪': 7, '杜': 7, '余': 7, '辛': 7, '谷': 7, '巫': 7, '狄': 7, '阮': 7, '冷': 7, '沙': 7, '言': 7, '車': 7, '岑': 7, '貝': 7, '辰': 7, '步': 7, '利': 7, '伯': 7, '伴': 7, '佟': 7, '別': 7, '判': 7, '君': 7, '吾': 7, '宏': 7, '局': 7,
        '林': 8, '周': 8, '邱': 8, '金': 8, '卓': 8, '季': 8, '岳': 8, '易': 8, '孟': 8, '邵': 8, '武': 8, '花': 8, '房': 8, '屈': 8, '宗': 8, '官': 8, '尚': 8, '祁': 8, '宓': 8, '明': 8, '杭': 8, '東': 8, '果': 8, '宛': 8, '孤': 8, '帛': 8, '忽': 8, '昔': 8, '松': 8, '牧': 8, '竺': 8, '長': 8,
        '侯': 9, '柯': 9, '洪': 9, '范': 9, '紀': 9, '姚': 9, '柳': 9, '段': 9, '韋': 9, '查': 9, '施': 9, '姜': 9, '柏': 9, '帥': 9, '俞': 9, '姬': 9, '冠': 9, '宣': 9, '郁': 9, '炳': 9, '柴': 9, '胡': 9, '相': 9, '皇': 9, '秋': 9, '紅': 9, '美': 9, '苗': 9, '計': 9, '軍': 9, '重': 9,
        '孫': 10, '高': 10, '翁': 10, '徐': 10, '唐': 10, '夏': 10, '涂': 10, '烏': 10, '殷': 10, '祝': 10, '桂': 10, '留': 10, '班': 10, '耿': 10, '馬': 10, '展': 10, '席': 10, '祖': 10, '索': 10, '袁': 10, '閃': 10, '倪': 10, '凌': 10, '宮': 10, '恩': 10, '時': 10, '晉': 10, '晏': 10, '真': 10, '神': 10, '秦': 10, '秘': 10, '能': 10, '貢': 10, '財': 10, '軒': 10, '郝': 10,
        '張': 11, '梁': 11, '許': 11, '郭': 11, '曹': 11, '崔': 11, '康': 11, '符': 11, '莊': 11, '麥': 11, '寇': 11, '浦': 11, '屠': 11, '婁': 11, '戚': 11, '區': 11, '培': 11, '婉': 11, '尉': 11, '粘': 11, '商': 11, '國': 11, '堅': 11, '堂': 11, '曼': 11, '梅': 11, '畢': 11, '章': 11, '莫': 11, '連': 11, '雪': 11, '頂': 11, '鹿': 11,
        '黃': 12, '彭': 12, '曾': 12, '游': 12, '賀': 12, '湯': 12, '馮': 12, '費': 12, '辜': 12, '項': 12, '覃': 12, '粟': 12, '單': 12, '斐': 12, '盛': 12, '寒': 12, '華': 12, '貴': 12, '買': 12, '雲': 12, '斯': 12, '喬': 12, '尊': 12, '就': 12, '悲': 12, '智': 12, '期': 12, '朝': 12, '焦': 12, '舒': 12, '童': 12, '開': 12, '閒': 12, '黑': 12,
        '楊': 13, '葉': 13, '董': 13, '詹': 13, '溫': 13, '萬': 13, '賈': 13, '楚': 13, '裘': 13, '祿': 13, '葛': 13, '塗': 13, '虞': 13, '雷': 13, '號': 13, '傳': 13, '微': 13, '愛': 13, '會': 13, '歲': 13, '義': 13, '農': 13, '路': 13, '載': 13, '鄒': 13,
        '廖': 14, '趙': 14, '齊': 14, '熊': 14, '管': 14, '蒲': 14, '臧': 14, '壽': 14, '裴': 14, '褚': 14, '翟': 14, '夢': 14, '實': 14, '榮': 14, '福': 14, '端': 14, '網': 14, '臺': 14, '蒙': 14, '語': 14,
        '劉': 15, '潘': 15, '鄧': 15, '黎': 15, '鈕': 15, '樊': 15, '樞': 15, '魯': 15, '墨': 15, '震': 15, '劍': 15, '增': 15, '德': 15, '慧': 15, '慶': 15, '滿': 15, '輝': 15, '銳': 15, '靠': 15, '蔣': 15,
        '陳': 16, '賴': 16, '盧': 16, '駱': 16, '穆': 16, '鮑': 16, '霍': 16, '錢': 16, '龍': 16, '學': 16, '憲': 16, '燕': 16, '糖': 16, '融': 16, '錦': 16, '靜': 16, '閻': 16,
        '謝': 17, '鍾': 17, '韓': 17, '戴': 17, '隋': 17, '聯': 17, '優': 17, '應': 17, '懂': 17, '戲': 17, '聰': 17, '薛': 17,
        '簡': 18, '魏': 18, '顏': 18, '豐': 18, '叢': 18, '歸': 18, '璧': 18, '雙': 18, '蕭': 18,
        '鄭': 19, '羅': 19, '蘇': 19, '龐': 19, '譚': 19, '關': 19, '麗': 19, '贊': 19, '鏡': 19, '願': 19,
        '嚴': 20, '藍': 20, '鐘': 20, '耀': 20, '黨': 20, '覺': 20,
        '顧': 21, '龔': 21, '鶴': 21,
        '權': 22, '藺': 22, '歡': 22
    }

    def get_stroke_count(name):
        if not isinstance(name, str) or not name.strip():
            return 999
        first_char = name.strip()[0]
        return stroke_map.get(first_char, 99) 

    processed_df["筆畫數"] = processed_df["執行醫生"].apply(get_stroke_count)
    processed_df = processed_df.sort_values(by=["筆畫數", "執行醫生"]).reset_index(drop=True)

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
    # 5. 網頁上方：明確顯示各個執行醫生的數量
    # ==========================================
    st.markdown("### 📈 執行醫師工作量統計")
    
    doctor_counts = processed_df["執行醫生"].value_counts()
    
    metric_cols = st.columns(len(doctor_counts)) if len(doctor_counts) > 0 else st.columns(1)
    for i, (doc, count) in enumerate(doctor_counts.items()):
        with metric_cols[i]:
            st.markdown(
                f"""
                <div style="background-color: rgba(240, 242, 246, 0.9); padding: 20px; border-radius: 10px; text-align: center; border: 1px solid #ddd;">
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
        "> - **執行醫師** 已自動依照姓氏筆畫由少到多排序。\n"
        "> - **病歷號** 已完整保留包含 `0` 開頭的所有數字。\n"
        "> - 診別若包含**「門」**字則顯示黑色；否則以 **紅色字體** 顯示。\n"
        "> - **執行日期** 已隱藏年份，並統一以 **紅色字體** 顯示。\n"
        "> - 若開單醫生**不在左側儲存的心臟科醫師名單內**，姓名會以 **紅色字體** 顯示並加上警示。"
    )

    def render_custom_table(df_data):
        html = '<div style="overflow-x: auto;">'
        html += '<style>'
        html += '.report-table { width: 100%; border-collapse: collapse; font-family: sans-serif; font-size: 14px; background-color: white; }'
        html += '.report-table th, .report-table td { border: 1px solid #ddd; padding: 8px 12px; text-align: center; }'
        html += '.report-table th { background-color: #f0f2f6; color: #31333F; }'
        html += '.red-text { color: red; font-weight: bold; }'
        html += '.warning-badge { background-color: #ff4b4b; color: white; padding: 2px 6px; border-radius: 4px; font-size: 12px; font-weight: bold; margin-left: 5px; }'
        html += '</style>'
        html += '<table class="report-table"><thead><tr>'
        html += '<th>項次</th><th>診別</th><th>病歷號</th><th>姓名</th><th>性別</th><th>執行日期</th><th>開單日期</th><th>開單醫生</th><th>執行醫生</th>'
        html += '</tr></thead><tbody>'

        for idx, row in df_data.iterrows():
            zbie_val = str(row['診別'])
            zbie_class = "" if "門" in zbie_val else "red-text"

            doc_display = row["開單醫生"]
            doc_class = ""
            if not row["是否心臟科"]:
                doc_class = "red-text"
                doc_display = f"{doc_display} <span class='warning-badge'>⚠️ 非心臟科</span>"

            html += f'<tr><td>{idx + 1}</td><td class="{zbie_class}">{zbie_val}</td><td>{row["病例號"]}</td><td>{row["姓名"]}</td><td>{row["性別"]}</td><td class="red-text">{row["執行日期"]}</td><td>{row["開單日期"]}</td><td class="{doc_class}">{doc_display}</td><td><b>{row["執行醫生"]}</b></td></tr>'

        html += '</tbody></table></div>'
        return html

    st.markdown(render_custom_table(processed_df), unsafe_allow_html=True)

    # ==========================================
    # 7. 匯出下載功能
    # ==========================================
    st.markdown("---")
    st.markdown("### 💾 匯出處理後報表")

    export_df = processed_df[
        ["診別", "病例號", "姓名", "性別", "執行日期", "開單日期", "開單醫生", "執行醫生"]
    ].copy()

    def to_excel(df_to_save):
      output = io.BytesIO()
      with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df_to_save.to_excel(writer, index=False, sheet_name="處理後報表")
        
        workbook = writer.book
        worksheet = writer.sheets["處理後報表"]
        
        red_format = workbook.add_format({'font_color': 'red'})
        
        zbie_col_idx = df_to_save.columns.get_loc("診別")
        date_col_idx = df_to_save.columns.get_loc("執行日期")
        doc_col_idx = df_to_save.columns.get_loc("開單醫生")
        
        for row_idx in range(len(df_to_save)):
            excel_row = row_idx + 1 
            
            zbie_val = str(df_to_save.iloc[row_idx, zbie_col_idx])
            if "門" not in zbie_val:
                worksheet.write_string(excel_row, zbie_col_idx, zbie_val, red_format)
                
            date_val = str(df_to_save.iloc[row_idx, date_col_idx])
            worksheet.write_string(excel_row, date_col_idx, date_val, red_format)
            
            doc_val = str(df_to_save.iloc[row_idx, doc_col_idx])
            if doc_val not in cardio_doctors:
                worksheet.write_string(excel_row, doc_col_idx, doc_val, red_format)
            
      return output.getvalue()

    excel_data = to_excel(export_df)

    st.download_button(
        label="📥 下載處理完成的 Excel 報表",
        data=excel_data,
        file_name="processed_schedule_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

  except Exception as e:
    st.error(f"❌ 讀取或處理檔案時發生錯誤：{e}")
    st.info("請確認上傳的是正確格式的醫療排程報表 Excel 檔案。")
else:
  st.info("💡 請由上方按鈕上傳您的排程報表檔案 (.xls / .xlsx) 來開始使用。")
