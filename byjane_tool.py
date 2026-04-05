import streamlit as st
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
import io

# --- 核心邏輯函數 ---
def get_vaule(col, row):
    return get_column_letter(col) + str(row)

def f_prod_pack_col(prod_pack):
    mapping = {
        "提袋": 20, "禮盒包裝（滿八入）": 21,
        "ByJane 馬年限定禮盒 2026（宅配）": 1, "ByJane 馬年限定禮盒 2026 （自取）": 1
    }
    return mapping.get(prod_pack, 0)

def f_prod_type_col(prod_type):
    mapping = {
        "Ａ ｜ The Medley Box  綜合風味": 3, "Ｂ｜ Waffle Lovers  人氣精選": 4,
        "Ｃ｜ The Refined Collection 成熟風味": 5, "Ｄ": 6,
        "經典原味": 7, "肉桂": 8, "濃心可可": 9, "藍莓乳酪": 10,
        "糖漬檸檬乳酪": 11, "芝麻": 12, "伯爵茶麻糬": 13, "抹茶紅豆麻糬": 14,
        "焙茶": 15, "培根楓糖起司": 16, "烤地瓜": 17, "焦糖杏仁奶油": 18,
        "鹽之花開心果可可": 19
    }
    return mapping.get(prod_type, 0)

# --- Streamlit 介面 ---
st.set_page_config(page_title="ByJane 訂單自動處理", layout="wide")
st.title("🧇 ByJane 訂單自動處理系統")
st.write("上傳『訂單報表.xlsx』後，系統將自動生成宅配明細、黑貓匯入檔及宅轉店檔案。")

uploaded_file = st.file_uploader("選擇 Excel 檔案", type=["xlsx"])

if uploaded_file:
    try:
        # 1. 讀取來源檔案
        wb_source = load_workbook(uploaded_file, data_only=True)
        ws = wb_source.active
        
        # --- 自動建立標題地圖 ---
        header_map = {}
        for col in range(1, ws.max_column + 1):
            title = ws.cell(row=1, column=col).value
            if title:
                header_map[str(title).strip()] = col

        # --- 欄位定義 (若找不到則使用預設索引) ---
        # 如果你的 Excel 標題名稱不同，請修改下方的字串
        col_id = header_map.get('訂單編號', 1)
        col_name = header_map.get('收件人姓名', 2)
        col_mobile = header_map.get('收件人手機', 5)
        col_addr = header_map.get('收件人地址', 7)
        col
