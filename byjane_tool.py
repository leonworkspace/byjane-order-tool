import streamlit as st
import pandas as pd
from openpyxl import Workbook, load_workbook
import io

# --- 核心邏輯函數 ---
def f_prod_pack_col(value):
    mapping = {
        "提袋": 20, "禮盒包裝（滿八入）": 21,
        "ByJane 馬年限定禮盒 2026（宅配）": 1, "ByJane 馬年限定禮盒 2026 （自取）": 1
    }
    return mapping.get(str(value).strip(), 0) if value else 0

def f_prod_type_col(value):
    mapping = {
        "Ａ ｜ The Medley Box  綜合風味": 3, "Ｂ｜ Waffle Lovers  人氣精選": 4,
        "Ｃ｜ The Refined Collection 成熟風味": 5, "Ｄ": 6,
        "經典原味": 7, "肉桂": 8, "濃心可可": 9, "藍莓乳酪": 10,
        "糖漬檸檬乳酪": 11, "芝麻": 12, "伯爵茶麻糬": 13, "抹茶紅豆麻糬": 14,
        "焙茶": 15, "培根楓糖起司": 16, "烤地瓜": 17, "焦糖杏仁奶油": 18,
        "鹽之花開心果可可": 19
    }
    return mapping.get(str(value).strip(), 0) if value else 0

# --- Streamlit 介面 ---
st.set_page_config(page_title="ByJane 訂單自動處理", layout="wide")
st.title("🧇 ByJane 訂單自動處理系統")

uploaded_file = st.file_uploader("請上傳『訂單報表.xlsx』", type=["xlsx"])

if uploaded_file:
    wb_source = load_workbook(uploaded_file, data_only=True)
    ws = wb_source.active
    
    # --- 1. 建立動態標題地圖 ---
    header_map = {}
    for col in range(1, ws.max_column + 1):
        title = ws.cell(row=1, column=col).value
        if title:
            header_map[str(title).strip()] = col

    # --- 2. 欄位定義 (根據你的要求調整預設索引) ---
    col_id      = header_map.get('訂單編號', 1)
    col_name    = header_map.get('收件人姓名', 2)
    col_mobile  = header_map.get('收件人手機', 5)
    col_addr    = header_map.get('收件人地址', 7)
    
    # 名稱預設 9, 款式預設 10, 數量預設 11
    col_pack    = header_map.get('商品名稱', 9) 
    col_type    = header_map.get('商品款式', 10) 
    col_val     = header_map.get('數量', 11)
    col_deliver = header_map.get('配送方式', 14)

    # --- 3. 初始化目標表格 ---
    wb_byjane = Workbook(); ws_byjane = wb_byjane.active
    ws_byjane.append(["訂單編號", "姓名", "A", "B", "C", "D", "原味", "肉桂", "可可", "藍莓", "檸檬", "芝麻", "伯爵", "抹茶", "焙茶", "培根", "地瓜", "焦糖", "開心果", "提袋", "禮盒", "總數"])

    wb_cat = Workbook(); ws_cat = wb_cat.active
    ws_cat.append(["收件人姓名", "收件人電話", "收件人手機", "收件人地址", "代收金額或到付", "件數", "品名(詳參數表)", "備註", "訂單編號", "希望配達時間(詳參數表)", "出貨日期(YYYY/MM/DD)", "預定配達日期(YYYY/MM/DD)", "溫層(詳參數表)", "尺寸(詳參數表)", "寄件人姓名", "寄件人電話", "寄件人手機", "寄件人地址", "保值金額", "品名說明", "是否列印(Y/N)", "是否捐贈(Y/N)", "統一編號", "手機載具", "愛心碼", "可刷卡(Y/N)", "手機支付(Y/N)"])

    wb_711 = Workbook(); ws_711 = wb_711.active
    ws_711.append(["訂單編號", "收件人姓名(必填)", "收件人手機(必填)", "FB名稱", "訂單備註", "代收金額", "門市編號(必填)", "匯款帳戶後五碼", "列印張數", "溫層(冷凍：0003)"])

    # --- 4. 數據處理 ---
    row_source = 2
    row_byjane = 1
    order_num_flag = 0
    prod_sum = 0
    
    while True:
        order_num_current = ws.cell(row=row_source, column=col_id).value
        order_num_next = ws.cell(row=row_source + 1, column=col_id).value
        
        if order_num_flag == 0:
            order_name_current = ws.cell(row=row_source, column=col_name).value
            row_byjane += 1
            ws_byjane.cell(row=row_byjane, column=1).value = order_num_current
            ws_byjane.cell(row=row_byjane, column=2).value = order_name_current
            order_num_flag = 1

        v_pack = ws.cell(row=row_source, column=col_pack).value
        v_type = ws.cell(row=row_source, column=col_type).value
        p_val  = ws.cell(row=row_source, column=col_val).value or 0
        
        # 處理包裝
        p_col = f_prod_pack_col(v_pack)
        if p_col != 0:
            ws_byjane.cell(row=row_byjane, column=p_col).value = p_val
            if p_col < 7: prod_sum += (p_val * 8)
        
        # 處理款式
        t_col = f_prod_type_col(v_type)
        if t_col != 0:
            ws_byjane.cell(row=row_byjane, column=t_col).value = p_val
            prod_sum += p_val

        # 訂單結束處理
        if order_num_current != order_num_next:
            ws_byjane.cell(row=row_byjane, column=22).value = prod_sum
            deliver = ws.cell(row=row_source, column=col_deliver).value
            mobile = ws.cell(row=row_source, column=col_mobile).value
            address = ws.cell(row=row_source, column=col_addr).value
            
            boxes = (prod_sum // 90) + (1 if prod_sum % 90 != 0 else 0) if prod_sum > 0 else 1

            if deliver == "黑貓冷凍宅配":
                cat_row = [""] * 27
                cat_row[0], cat_row[1], cat_row[2], cat_row[3] = order_name_current, mobile, mobile, address
                cat_row[5], cat_row[6], cat_row[8] = boxes, 1, order_num_current
                cat_row[9], cat_row[12], cat_row[13] = 4, 3, 1
                cat_row[14], cat_row[15], cat_row[16], cat_row[17] = "Byjane簡", "0960-319-998", "0960-319-998", "台南市中西區西賢五街26號"
                ws_cat.append(cat_row)
            elif deliver and ("711" in str(deliver) or "快速到店" in str(deliver)):
                ws_711.append([order_num_current, order_name_current, mobile, "", order_num_current, "", "", "", boxes, "0003"])

            prod_sum = 0
            order_num_flag = 0

        if order_num_next is None: break
        row_source += 1

    st.success("✨ 處理完成！")
    def get_io(wb):
        out = io.BytesIO(); wb.save(out); return out.getvalue()
    c1, c2, c3 = st.columns(3)
    c1.download_button("📂 下載宅配明細", get_io(wb_byjane), "宅配明細.xlsx")
    c2.download_button("🚚 下載黑貓單", get_io(wb_cat), "黑貓單.xlsx")
    c3.download_button("🏪 下載宅轉店", get_io(wb_711), "宅轉店.xlsx")
