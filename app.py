import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from datetime import datetime
import collections
import io

# --------------------
# 模組 1: 解合併並填入原值
# --------------------
def unmerge_and_fill(ws):
    merged_ranges = list(ws.merged_cells.ranges)
    for merged_range in merged_ranges:
        cells = list(ws[merged_range.coord])
        value = cells[0][0].value
        ws.unmerge_cells(str(merged_range))
        for row in cells:
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
                        shift_type = str(ws.cell(i, c).value).strip() if ws.cell(i, c).value else ""
                        if isinstance(ws.cell(i, c).value, datetime) or shift_type == "":
                            break
                        if shift_type in ["早","午","晚"]:
                            i += 1
                            while i <= max_row:
                                val_cell = ws.cell(i, c).value
                                if isinstance(val_cell, datetime):
                                    break
                                val = str(val_cell).strip() if val_cell else ""
                                if val in ["早","午","晚"] or val == "":
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
    if not emp_title:
        return ""
    if emp_title == "醫師":
        class_code = "★醫師★"
    else:
        class_code = "【員工】"
    class_code += shift_type + "班"
    return class_code

def create_shift_analysis(wb, df_consolidated, ws_employee):
    emp_dict = {}
    for row in ws_employee.iter_rows(min_row=2, values_only=True):
        emp_id = row[0] if len(row) > 0 else ""
        name = row[1] if len(row) > 1 else ""
        dept = row[2] if len(row) > 2 else ""
        title = row[3] if len(row) > 3 else ""
        if name:
            emp_dict[name.strip()] = (emp_id, dept, title)

    shift_dict = {}
    for idx, row in df_consolidated.iterrows():
        clinic = row["診所"]
        date_str = row["日期"]
        shift_type = row["班別"]
        name = row["姓名"]
        e_value = row["A欄資料"]
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
# 模組 4: 建立班別總表
# --------------------
def create_shift_summary(wb):
    ws_analysis = wb["班別分析"]
    all_dates = [datetime(2025,8,i).strftime("%Y-%m-%d") for i in range(1,32)]
    shift_dict = collections.defaultdict(dict)

    for row in ws_analysis.iter_rows(min_row=2, values_only=True):
        emp_id = row[1] if len(row) > 1 else ""
        emp_name = row[2] if len(row) > 2 else ""
        shift_date = row[6] if len(row) > 6 else ""
        class_code = row[8] if len(row) > 8 else ""
        if not emp_id or not emp_name or not shift_date:
            continue
        emp_key = f"{emp_id}|{emp_name}"
        shift_dict[emp_key][shift_date] = class_code

    if "班別總表" in wb.sheetnames:
        ws_target = wb["班別總表"]
        ws_target.delete_rows(1, ws_target.max_row)
    else:
        ws_target = wb.create_sheet("班別總表")

    ws_target.append(["員工編號","員工姓名"] + all_dates)
    for emp_key, date_map in shift_dict.items():
        emp_id, emp_name = emp_key.split("|")
        row = [emp_id, emp_name] + [date_map.get(d,"") for d in all_dates]
        ws_target.append(row)

# --------------------
# Streamlit 主程式
# --------------------
st.title("班表處理系統")

# 上傳班表 Excel
shift_file = st.file_uploader("上傳班表 Excel", type=["xlsx","xlsm"])
if shift_file:
    wb_shift = load_workbook(shift_file)
    excluded = ["彙整結果","班別分析","班別總表"]
    selectable_sheets = [s for s in wb_shift.sheetnames if s not in excluded]
    selected_sheets = st.multiselect("選擇要處理的工作表", selectable_sheets)

    # 上傳員工資料
    employee_file = st.file_uploader("上傳員工資料 Excel", type=["xlsx","xlsm"])
    if employee_file:
        wb_employee = load_workbook(employee_file)
        employee_sheet_name = st.selectbox("選擇員工資料工作表", wb_employee.sheetnames)
        ws_employee = wb_employee[employee_sheet_name]

        if st.button("執行班表處理"):
            df_consolidated = consolidate_selected_sheets(wb_shift, selected_sheets)
            create_shift_analysis(wb_shift, df_consolidated, ws_employee)
            create_shift_summary(wb_shift)

            # 儲存到 BytesIO，提供下載
            output_stream = io.BytesIO()
            wb_shift.save(output_stream)
            output_stream.seek(0)
            st.success("班表處理完成！")
            st.download_button("下載結果 Excel", data=output_stream, file_name="output.xlsx")
