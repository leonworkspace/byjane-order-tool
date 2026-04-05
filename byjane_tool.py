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
st.title("🧇 ByJane 訂單自動處理系統 (標題自動辨識版)")

uploaded_file = st.file_uploader("請上傳『訂單報表.xlsx』", type=["xlsx"])

if uploaded_file:
    wb_source = load_workbook(uploaded_file)
    ws = wb_source.active
    
    # --- 1. 自動建立標題地圖 (核心更新) ---
    header_map = {}
    for col in range(1, ws.max_column + 1):
        title = ws.cell(row=1, column=col).value
        if title:
            header_map[title.strip()] = col # 去除空格確保精準度

    # 檢查必要欄位是否存在 (你可以根據實際 Excel 標題微調關鍵字)
    try:
        COL_ORDER_ID = header_map['訂單編號']
        COL_NAME = header_map['收件人姓名']
        COL_MOBILE = header_map['收件人手機']
        COL_ADDR = header_map['收件人地址']
        COL_PROD_PACK = header_map['包裝'] # 請確認你的 Excel 標題是否為這兩個字
        COL_PROD_TYPE = header_map['品項名稱']
        COL_PROD_VAL = header_map['數量']
        COL_DELIVER = header_map['配送方式']
    except KeyError as e:
        st.error(f"❌ 找不到必要欄位：{e}。請檢查 Excel 第一列標題名稱是否正確。")
        st.stop()

    # --- 2. 初始化目標檔案 (保持原樣) ---
    wb_byjane = Workbook()
    ws_byjane = wb_byjane.active
    ws_byjane.append(["訂單編號", "姓名", "綜合", "人氣", "成熟", "D", "原味", "肉桂", "可可", "藍莓", "檸檬", "芝麻", "伯爵", "抹茶", "焙茶", "培根", "地瓜", "焦糖", "開心果", "提袋", "禮盒", "總數"])

    wb_cat = Workbook()
    ws_cat = wb_cat.active
    ws_cat.append(["收件人姓名", "收件人電話", "收件人手機", "收件人地址", "代收金額或到付", "件數", "品名(詳參數表)", "備註", "訂單編號", "希望配達時間(詳參數表)", "出貨日期(YYYY/MM/DD)", "預定配達日期(YYYY/MM/DD)", "溫層(詳參數表)", "尺寸(詳參數表)", "寄件人姓名", "寄件人電話", "寄件人手機", "寄件人地址", "保值金額", "品名說明", "是否列印(Y/N)", "是否捐贈(Y/N)", "統一編號", "手機載具", "愛心碼", "可
