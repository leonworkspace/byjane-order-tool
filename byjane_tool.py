import streamlit as st
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
import io

# --- 核心邏輯函數 (保留您原本的 Mapping) ---
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
st.set_page_config(page_title="ByJane 訂單自動轉換", layout="wide")
st.title("🧇 ByJane 訂單自動處理系統")

uploaded_file = st.file_uploader("請上傳『訂單報表.xlsx』", type=["xlsx"])

if uploaded_file:
    # 1. 讀取來源
    wb_source = load_workbook(uploaded_file)
    ws = wb_source.active
    
    # 2. 初始化目標檔案
    # A. 宅配明細 (自定義標題)
    wb_byjane = Workbook()
    ws_byjane = wb_byjane.active
    ws_byjane.append(["訂單編號", "姓名", "綜合", "人氣", "成熟", "D", "原味", "肉桂", "可可", "藍莓", "檸檬", "芝麻", "伯爵", "抹茶", "焙茶", "培根", "地瓜", "焦糖", "開心果", "提袋", "禮盒", "總數"])

    # B. 黑貓託運單 (27 欄標題)
    wb_cat = Workbook()
    ws_cat = wb_cat.active
    cat_headers = ["收件人姓名", "收件人電話", "收件人手機", "收件人地址", "代收金額或到付", "件數", "品名(詳參數表)", "備註", "訂單編號", "希望配達時間(詳參數表)", "出貨日期(YYYY/MM/DD)", "預定配達日期(YYYY/MM/DD)", "溫層(詳參數表)", "尺寸(詳參數表)", "寄件人姓名", "寄件人電話", "寄件人手機", "寄件人地址", "保值金額", "品名說明", "是否列印(Y/N)", "是否捐贈(Y/N)", "統一編號", "手機載具", "愛心碼", "可刷卡(Y/N)", "手機支付(Y/N)"]
    ws_cat.append(cat_headers)

    # C. 黑貓宅轉店 (10 欄標題)
    wb_711 = Workbook()
    ws_711 = wb_711.active
    ws_711.append(["訂單編號", "收件人姓名(必填)", "收件人手機(必填)", "FB名稱", "訂單備註", "代收金額", "門市編號(必填)", "匯款帳戶後五碼", "列印張數", "溫層(冷凍：0003)"])

    # 3. 處理邏輯
    order_info = ['', '', '', '', '', '', '', '', '', 4, '', '', 3, 1, 'Byjane簡', '0960-319-998', '0960-319-998', '台南市中西區西賢五街26號']
    
    row_source = 2
    row_byjane = 1
    order_num_flag = 0
    prod_sum = 0
    
    while True:
        order_num_next = ws[get_vaule(1, row_source + 1)].value
        
        if order_num_flag == 0:
            order_num_current = ws[get_vaule(1, row_source)].value
            order_name_current = ws[get_vaule(2, row_source)].value
            row_byjane += 1
            ws_byjane.cell(row=row_byjane, column=1).value = order_num_current
            ws_byjane.cell(row=row_byjane, column=2).value = order_name_current
            order_num_flag = 1

        # 產品分配邏輯
        p_pack = ws[get_vaule(10, row_source)].value
        p_type = ws[get_vaule(11, row_source)].value
        p_val = ws[get_vaule(12, row_source)].value or 0
        
        p_col = f_prod_pack_col(p_pack)
        if p_col != 0 and p_col != 1:
            ws_byjane.cell(row=row_byjane, column=p_col).value = p_val
        else:
            t_col = f_prod_type_col(p_type)
            if t_col != 0:
                ws_byjane.cell(row=row_byjane, column=t_col).value = p_val
                prod_sum += (p_val * 8 if t_col < 7 else p_val)

        # 訂單結束判斷與拆表
        if order_num_current != order_num_next:
            ws_byjane.cell(row=row_byjane, column=22).value = prod_sum
            
            deliver = ws[get_vaule(14, row_source)].value
            mobile = ws[get_vaule(5, row_source)].value
            address = ws[get_vaule(7, row_source)].value
            
            # 計算箱數
            boxes = (prod_sum // 90) + (1 if prod_sum % 90 != 0 else 0) if prod_sum > 0 else 1

            if deliver == "黑貓冷凍宅配":
                new_row = [order_name_current, mobile, mobile, address, "", boxes, 1, "", order_num_current, 4, "", "", 3, 1, "Byjane簡", "0960-319-998", "0960-319-998", "台南市中西區西賢五街26號"]
                ws_cat.append(new_row + [""] * (27 - len(new_row)))

            elif "711" in str(deliver) or "快速到店" in str(deliver):
                ws_711.append([order_num_current, order_name_current, mobile, "", order_num_current, "", "", "", boxes, "0003"])

            prod_sum = 0
            order_num_flag = 0

        if order_num_next is None: break
        row_source += 1

    # 4. 下載按鈕區
    st.success("✨ 處理完畢！")
    c1, c2, c3 = st.columns(3)
    
    def get_io(wb):
        out = io.BytesIO()
        wb.save(out)
        return out.getvalue()

    c1.download_button("📂 下載宅配明細", get_io(wb_byjane), "宅配明細.xlsx")
    c2.download_button("🚚 下載黑貓單", get_io(wb_cat), "黑貓匯入檔.xlsx")
    c3.download_button("🏪 下載宅轉店", get_io(wb_711), "宅轉店匯入檔.xlsx")
