import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import collections
from datetime import datetime
import tempfile
import os

# --------------------
# 模組 1: 解合併並填入原值
# --------------------
def unmerge_and_fill(ws):
    for merged in list(ws.merged_cells.ranges):
        value = ws.cell(merged.min_row, merged.min_col).value
        ws.unmerge_cells(str(merged))
        for row in ws[merged.coord]:
            for cell in row:
                cell.value = value

# --------------------
# 模組 2: 彙整班表資料
# --------------------
def consolidate_selected_sheets(wb, sheet_names):
    all_data = []
    for sheet_name in sheet_names:
        ws = wb[sheet_name]
        unmerge_and_fill(ws)
        clinic_name = str(ws.cell(row=1, column=1).value)[:4]
        max_row = ws.max_row
        max_col = ws.max_column
        for r in range(1, max_row+1):
            for c in range(2, max_col+1):
                cell_value = ws.cell(r, c).value
                if isinstance(cell_value, datetime):
                    date_val = cell_value
                    i = r + 3
                    while i <= max_row:
                        shift_type = str(ws.cell(i, c).value).strip()
                        if isinstance(ws.cell(i, c).value, datetime) or shift_type == "":
                            break
                        if shift_type in ["早","午","晚"]:
                            i += 1
                            while i <= max_row:
                                if isinstance(ws.cell(i, c).value, datetime):
                                    break
                                val = str(ws.cell(i, c).value).strip()
                                if val in ["早","午","晚"]:
                                    break
                                all_data.append([
                                    clinic_name,
                                    date_val.strftime("%Y/%m/%d"),
                                    shift_type,
                                    val,
                                    ws.cell(i, 1).value,
                                    ws.cell(i, 21).value
                                ])
                                i += 1
                            i -= 1
                        i += 1
    df = pd.DataFrame(all_data, columns=["診所","日期","班別","姓名","A欄資料","U欄資料"])
    return df

# --------------------
# 模組 3: 建立班別分析表
# --------------------
def format_shift_order(shift_str):
    result = ""
    for s in ["早","午","晚"]:
        if s in shift_str:
            result += s
    return result

def get_class_code(emp_title, clinic_name, shift_type):
    """
    根據員工職稱、診所名稱和班別類型生成班別代碼
    """
    if not emp_title or emp_title.strip() == "":
        return ""
    
    emp_title = emp_title.strip()
    
    # 處理純早班的特殊職稱
    pure_morning_titles = ["早班護理師", "早班內視鏡助理", "醫務專員", "兼職早班內視鏡助理"]
    if emp_title in pure_morning_titles:
        return "【員工】純早班"
    
    # 根據職稱決定基本代碼
    class_code = ""
    if emp_title == "醫師":
        class_code = "★醫師★"
    elif emp_title in ["櫃臺", "護理師", "兼職護理師", "兼職跟診助理", "副店長"]:
        class_code = "【員工】"
    else:
        # 處理其他特殊情況
        if "副店長" in emp_title:
            class_code = "【員工】"
        elif "店長" in emp_title or "護士" in emp_title:
            class_code = "◇主管◇"
        else:
            class_code = ""
    
    # 如果不是早班，根據診所名稱添加地區代碼
    if shift_type != "早":
        target_clinics = ["上吉診所", "立吉診所", "上承診所", "立全診所", "立竹診所", "立順診所", "上京診所"]
        if clinic_name in target_clinics:
            class_code += "板土中京"
        elif clinic_name == "立丞診所":
            class_code += "立丞"
    
    # 根據班別類型添加班別後綴
    shift_suffix_map = {
        "早": "早班",
        "午晚": "午晚班", 
        "早午晚": "全天班",
        "早晚": "早晚班",
        "午": "午班",
        "晚": "晚班",
        "早午": "早午班"
    }
    
    if shift_type in shift_suffix_map:
        class_code += shift_suffix_map[shift_type]
    
    # 修正重複的早班問題
    if class_code.endswith("早班早班"):
        class_code = class_code.replace("早班早班", "早班")
    
    return class_code

def create_shift_analysis(wb, df_consolidated, ws_employee):
    emp_dict = {}
    for row in ws_employee.iter_rows(min_row=2, values_only=True):
        emp_id, name, dept, title = row[:4]
        if name:
            emp_dict[name.strip()] = (emp_id, dept, title)

    shift_dict = {}
    for idx, row in df_consolidated.iterrows():
        clinic, date_str, shift_type, name, e_value, _ = row
        if not name or len(name) > 4:
            continue
        key = f"{name}|{date_str}|{clinic}|{e_value}"
        if key not in shift_dict:
            shift_dict[key] = shift_type
        else:
            shift_dict[key] += " " + shift_type

    if "班別分析" in wb.sheetnames:
        ws_target = wb["班別分析"]
        ws_target.delete_rows(1, ws_target.max_row)
    else:
        ws_target = wb.create_sheet("班別分析")

    headers = ["診所","員工編號","所屬部門","姓名","職稱","日期","班別","E欄資料","班別代碼"]
    for col_idx, h in enumerate(headers, 1):
        ws_target.cell(row=1, column=col_idx, value=h)

    for key, shift_types in shift_dict.items():
        name, date_str, clinic, e_value = key.split("|")
        shift_type = format_shift_order(shift_types)
        emp_info = emp_dict.get(name, ("","",""))
        emp_id, emp_dept, emp_title = emp_info
        ws_target.append([clinic, emp_id, emp_dept, name, emp_title, date_str, shift_type, e_value, get_class_code(emp_title, clinic, shift_type)])

# --------------------
# 模組 4: 建立班別總表（修正版）
# --------------------
def create_shift_summary(wb):
    ws_analysis = wb["班別分析"]
    
    # 修正：從實際資料中取得日期範圍，並轉換為文字格式
    all_dates = set()
    for row in ws_analysis.iter_rows(min_row=2, values_only=True):
        if len(row) >= 6 and row[5]:  # 第6欄是日期
            # 確保日期是文字格式
            date_value = row[5]
            if isinstance(date_value, datetime):
                date_str = date_value.strftime("%Y/%m/%d")
            else:
                date_str = str(date_value)
            all_dates.add(date_str)
    
    all_dates = sorted(list(all_dates))
    shift_dict = collections.defaultdict(dict)

    for row in ws_analysis.iter_rows(min_row=2, values_only=True):
        if len(row) < 9:  # 確保有足夠的欄位
            continue
        _, emp_id, emp_dept, emp_name, emp_title, shift_date, shift_type, e_value, class_code = row[:9]
        
        if not emp_id or not emp_name or not shift_date:
            continue
        
        # 確保日期是文字格式
        if isinstance(shift_date, datetime):
            shift_date_str = shift_date.strftime("%Y/%m/%d")
        else:
            shift_date_str = str(shift_date)
            
        emp_key = f"{emp_id}|{emp_name}"
        shift_dict[emp_key][shift_date_str] = class_code

    if "班別總表" in wb.sheetnames:
        ws_target = wb["班別總表"]
        ws_target.delete_rows(1, ws_target.max_row)
    else:
        ws_target = wb.create_sheet("班別總表")

    # 寫入標題行，確保日期是文字格式
    headers = ["員工編號", "員工姓名"] + all_dates
    for col_idx, header in enumerate(headers, 1):
        cell = ws_target.cell(row=1, column=col_idx, value=str(header))
        # 設定為文字格式
        cell.number_format = '@'
    
    # 寫入資料行
    for emp_key, date_map in shift_dict.items():
        emp_id, emp_name = emp_key.split("|")
        row_data = [str(emp_id), str(emp_name)] + [date_map.get(d,"") for d in all_dates]
        
        # 逐一寫入每個儲存格並設定為文字格式
        row_num = ws_target.max_row + 1
        for col_idx, value in enumerate(row_data, 1):
            cell = ws_target.cell(row=row_num, column=col_idx, value=str(value))
            # 對日期欄位設定文字格式
            if col_idx > 2:  # 日期欄位從第3欄開始
                cell.number_format = '@'

# --------------------
# Streamlit 主程式（修正版）
# --------------------
st.title("班表處理器")
shift_file = st.file_uploader("上傳班表 Excel 檔案", type=["xlsx","xlsm"])
employee_file = st.file_uploader("上傳員工資料 Excel 檔案", type=["xlsx","xlsm"])

if shift_file and employee_file:
    try:
        wb_shift = load_workbook(shift_file)
        wb_employee = load_workbook(employee_file)

        selectable_sheets = [s for s in wb_shift.sheetnames if s not in ["彙整結果","班別分析","班別總表"]]
        selected_sheets = st.multiselect("選擇要處理的工作表", selectable_sheets)
        
        if not selected_sheets:
            st.warning("請至少選擇一個工作表")
        else:
            employee_sheet_name = st.selectbox("選擇員工資料工作表", wb_employee.sheetnames)
            
            if employee_sheet_name:
                ws_employee = wb_employee[employee_sheet_name]

                if st.button("開始處理"):
                    with st.spinner("正在處理班表資料..."):
                        # 彙整班表資料
                        df_consolidated = consolidate_selected_sheets(wb_shift, selected_sheets)
                        
                        # 顯示彙整結果預覽
                        st.subheader("彙整結果預覽")
                        st.dataframe(df_consolidated.head(10))
                        
                        # 建立班別分析表（用於生成班別總表的中間步驟）
                        create_shift_analysis(wb_shift, df_consolidated, ws_employee)
                        
                        # 建立班別總表
                        create_shift_summary(wb_shift)

                        # 創建新的工作簿，只包含班別總表
                        from openpyxl import Workbook
                        new_wb = Workbook()
                        new_ws = new_wb.active
                        new_ws.title = "班別總表"
                        
                        # 複製班別總表的內容到新工作簿，確保日期為文字格式
                        ws_summary = wb_shift["班別總表"]
                        for row_idx, row in enumerate(ws_summary.iter_rows(values_only=True), 1):
                            for col_idx, value in enumerate(row, 1):
                                cell = new_ws.cell(row=row_idx, column=col_idx, value=str(value) if value is not None else "")
                                # 設定所有儲存格為文字格式
                                cell.number_format = '@'
                        
                        # 儲存到暫存檔
                        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_file:
                            new_wb.save(tmp_file.name)
                            
                            # 讀取檔案內容
                            with open(tmp_file.name, "rb") as f:
                                file_data = f.read()
                            
                            # 清理暫存檔
                            os.unlink(tmp_file.name)
                            
                        st.success("班別總表已生成完成！")
                        st.download_button(
                            label="下載班別總表 Excel",
                            data=file_data,
                            file_name="班別總表.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                        
                        # 顯示處理統計
                        st.subheader("處理統計")
                        st.write(f"總共處理了 {len(df_consolidated)} 筆班表記錄")
                        st.write(f"處理了 {len(selected_sheets)} 個工作表")
                        
                        # 顯示班別總表預覽
                        st.subheader("班別總表預覽")
                        summary_data = []
                        for row in ws_summary.iter_rows(values_only=True):
                            summary_data.append(row)
                        
                        if summary_data:
                            df_summary = pd.DataFrame(summary_data[1:], columns=summary_data[0])
                            st.dataframe(df_summary.head(10))
                        
    except Exception as e:
        st.error(f"處理過程中發生錯誤：{str(e)}")
        st.write("請檢查檔案格式是否正確")
else:
    st.info("請上傳班表 Excel 檔案和員工資料 Excel 檔案以開始處理")
