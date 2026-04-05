import streamlit as st
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill
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
st.title(" waffle ByJane 訂單自動處理系統")

uploaded_file = st.file_uploader("請上傳『訂單報表.xlsx』", type=["xlsx"])

if uploaded_file:
    wb_source = load_workbook(uploaded_file, data_only=True)
    ws = wb_source.active
    
    header_map = {}
    for col in range(1, ws.max_column + 1):
        title = ws.cell(row=1, column=col).value
        if title: header_map[str(title).strip()] = col

    col_id = header_map.get('訂單編號', 1)
    col_name = header_map.get('收件人姓名', 2)
    col_mobile = header_map.get('收件人手機', 5)
    col_addr = header_map.get('收件人地址', 7)
    col_pay_stat = header_map.get('付款狀態', 7)
    col_pack = header_map.get('商品名稱', 9) 
    col_type = header_map.get('商品款式', 10) 
    col_val = header_map.get('數量', 11)
    col_deliver = header_map.get('配送方式', 14)

    pink_fill = PatternFill(start_color="FFC9CA", end_color="FFC9CA", fill_type="solid")

    # --- 數據處理 (第一階段：收集所有資料並統計有效欄位) ---
    all_data = []
    active_cols = set() # 記錄哪些品項欄位有數字
    all_titles = ["訂單編號", "姓名", "A", "B", "C", "D", "原味", "肉桂", "可可", "藍莓", "檸檬", "芝麻", "伯爵", "抹茶", "焙茶", "培根", "地瓜", "焦糖", "開心果", "提袋", "禮盒"]
    
    row_source = 2
    order_num_flag = 0
    current_order_data = {}
    
    while True:
        order_num_current = ws.cell(row=row_source, column=col_id).value
        order_num_next = ws.cell(row=row_source + 1, column=col_id).value
        
        if order_num_flag == 0:
            current_order_data = {
                "id": order_num_current,
                "name": ws.cell(row=row_source, column=col_name).value,
                "unpaid": (str(ws.cell(row=row_source, column=col_pay_stat).value).strip() == "等待付款"),
                "items": {},
                "sum": 0,
                "deliver": ws.cell(row=row_source, column=col_deliver).value,
                "mobile": ws.cell(row=row_source, column=col_mobile).value,
                "address": ws.cell(row=row_source, column=col_addr).value
            }
            order_num_flag = 1

        p_val = ws.cell(row=row_source, column=col_val).value or 0
        if p_val > 0:
            p_col = f_prod_pack_col(ws.cell(row=row_source, column=col_pack).value)
            t_col = f_prod_type_col(ws.cell(row=row_source, column=col_type).value)
            
            target_col = p_col if p_col != 0 else t_col
            if target_col != 0:
                current_order_data["items"][target_col] = current_order_data["items"].get(target_col, 0) + p_val
                active_cols.add(target_col)
                current_order_data["sum"] += (p_val * 8 if target_col < 7 else p_val)

        if order_num_current != order_num_next:
            all_data.append(current_order_data)
            order_num_flag = 0

        if order_num_next is None: break
        row_source += 1

    # --- 第二階段：產出表格 ---
    # 決定要顯示哪些欄位 (1, 2 必顯，22 總數必顯，其餘看 active_cols)
    used_indices = sorted(list(active_cols))
    final_header = ["訂單編號", "姓名"] + [all_titles[i-1] for i in used_indices if i > 2] + ["總數"]

    wb_byjane = Workbook(); ws_byjane = wb_byjane.active
    ws_byjane.append(final_header)

    wb_cat = Workbook(); ws_cat = wb_cat.active
    ws_cat.append(["收件人姓名", "收件人電話", "收件人手機", "收件人地址", "代收金額或到付", "件數", "品名", "備註", "訂單編號", "希望配達時間", "出貨日期", "預定配達日期", "溫層", "尺寸", "寄件人姓名", "寄件人電話", "寄件人手機", "寄件人地址", "保值金額", "品名說明", "是否列印", "是否捐贈", "統一編號", "手機載具", "愛心碼", "可刷卡", "手機支付"])

    wb_711 = Workbook(); ws_711 = wb_711.active
    ws_711.append(["訂單編號", "收件人姓名(必填)", "收件人手機(必填)", "FB名稱", "訂單備註", "代收金額", "門市編號(必填)", "匯款帳戶後五碼", "列印張數", "溫層(冷凍：0003)"])

    for data in all_data:
        # A. 宅配明細 (動態填入)
        row_vals = [data["id"], data["name"]]
        for idx in used_indices:
            if idx > 2: row_vals.append(data["items"].get(idx, ""))
        row_vals.append(data["sum"] if data["sum"] > 0 else "")
        
        ws_byjane.append(row_vals)
        if data["unpaid"]:
            for c in range(1, len(row_vals) + 1):
                ws_byjane.cell(row=ws_byjane.max_row, column=c).fill = pink_fill

        # B. 黑貓 & 711 (邏輯不變)
        boxes = (data["sum"] // 90) + (1 if data["sum"] % 90 != 0 else 0) if data["sum"] > 0 else 1
        if data["deliver"] == "黑貓冷凍宅配":
            cat_row = [""] * 27
            cat_row[0], cat_row[1], cat_row[2], cat_row[3] = data["name"], data["mobile"], data["mobile"], data["address"]
            cat_row[5], cat_row[6], cat_row[8] = boxes, 1, data["id"]
            cat_row[9], cat_row[12], cat_row[13] = 4, 3, 1
            cat_row[14], cat_row[15], cat_row[16], cat_row[17] = "Byjane簡", "0960-319-998", "0960-319-998", "台南市中西區西賢五街26號"
            ws_cat.append(cat_row)
        elif data["deliver"] and ("711" in str(data["deliver"]) or "快速到店" in str(data["deliver"])):
            ws_711.append([data["id"], data["name"], data["mobile"], "", data["id"], "", "", "", boxes, "0003"])

    st.success("✨ 處理完成！")
    def get_io(wb):
        out = io.BytesIO(); wb.save(out); return out.getvalue()
    c1, c2, c3 = st.columns(3)
    c1.download_button("📂 宅配明細", get_io(wb_byjane), "宅配明細.xlsx")
    c2.download_button("🚚 黑貓單", get_io(wb_cat), "黑貓單.xlsx")
    c3.download_button("🏪 宅轉店", get_io(wb_711), "宅轉店.xlsx")
