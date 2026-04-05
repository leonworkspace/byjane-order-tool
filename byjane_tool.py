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

uploaded_file = st.file_uploader("選擇 Excel 檔案", type=["xlsx"])

if uploaded_file:
    # 1. 嘗試讀取檔案與建立地圖
    wb_source = load_workbook(uploaded_file, data_only=True)
    ws = wb_source.active
    
    header_map = {}
    for col in range(1, ws.max_column + 1):
        title = ws.cell(row=1, column=col).value
        if title:
            header_map[str(title).strip()] = col

    # 2. 定義欄位位置 (使用 try-except 包裹這段欄位檢核)
    try:
        col_id = header_map.get('訂單編號', 1)
        col_name = header_map.get('收件人姓名', 2)
        col_mobile = header_map.get('收件人手機', 5)
        col_addr = header_map.get('收件人地址', 7)
        col_pack = header_map.get('包裝', 10) 
        col_type = header_map.get('品項名稱', 11)
        col_val = header_map.get('數量', 12)
        col_deliver = header_map.get('配送方式', 14)
        
        if '訂單編號' not in header_map:
            st.warning("⚠️ 找不到『訂單編號』標題，將使用預設位置。")
            
    except Exception as e:
        st.error(f"解析標題時發生錯誤: {e}")
        st.stop()

    # 3. 初始化目標表格
    wb_byjane = Workbook()
    ws_byjane = wb_byjane.active
    ws_byjane.append(["訂單編號", "姓名", "綜合", "人氣", "成熟", "D", "原味", "肉桂", "可可", "藍莓", "檸檬", "芝麻", "伯爵", "抹茶", "焙茶", "培根", "地瓜", "焦糖", "開心果", "提袋", "禮盒", "總數"])

    wb_cat = Workbook()
    ws_cat = wb_cat.active
    ws_cat.append(["收件人姓名", "收件人電話", "收件人手機", "收件人地址", "代收金額或到付", "件數", "品名(詳參數表)", "備註", "訂單編號", "希望配達時間(詳參數表)", "出貨日期(YYYY/MM/DD)", "預定配達日期(YYYY/MM/DD)", "溫層(詳參數表)", "尺寸(詳參數表)", "寄件人姓名", "寄件人電話", "寄件人手機", "寄件人地址", "保值金額", "品名說明", "是否列印(Y/N)", "是否捐贈(Y/N)", "統一編號", "手機載具", "愛心碼", "可刷卡(Y/N)", "手機支付(Y/N)"])

    wb_711 = Workbook()
    ws_711 = wb_711.active
    ws_711.append(["訂單編號", "收件人姓名(必填)", "收件人手機(必填)", "FB名稱", "訂單備註", "代收金額", "門市編號(必填)", "匯款帳戶後五碼", "列印張數", "溫層(冷凍：0003)"])

    # 4. 處理數據
    row_source = 2
    row_byjane = 1
    order_num_flag = 0
    prod_sum = 0
    
    with st.spinner('正在分析訂單數據...'):
        while True:
            order_num_next = ws.cell(row=row_source + 1, column=col_id).value
            
            if order_num_flag == 0:
                order_num_current = ws.cell(row=row_source, column=col_id).value
                order_name_current = ws.cell(row=row_source, column=col_name).value
                row_byjane += 1
                ws_byjane.cell(row=row_byjane, column=1).value = order_num_current
                ws_byjane.cell(row=row_byjane, column=2).value = order_name_current
                order_num_flag = 1

            p_pack = ws.cell(row=row_source, column=col_pack).value
            p_type = ws.cell(row=row_source, column=col_type).value
            p_val = ws.cell(row=row_source, column=col_val).value or 0
            
            p_col = f_prod_pack_col(p_pack)
            if p_col != 0 and p_col != 1:
                ws_byjane.cell(row=row_byjane, column=p_col).value = p_val
            else:
                t_col = f_prod_type_col(p_type)
                if t_col != 0:
                    ws_byjane.cell(row=row_byjane, column=t_col).value = p_val
                    prod_sum += (p_val * 8 if t_col
