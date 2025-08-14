import streamlit as st
import pandas as pd
from openpyxl import load_workbook, Workbook
from io import BytesIO
from datetime import datetime
import collections

# --------------------
# 輔助函數
# --------------------
def unmerge_and_fill(ws):
    for row in ws.iter_rows():
        for cell in row:
            if cell.merge_cells:
                merged_range = cell.merged_cells.ranges[0]
                value = cell.value
                ws.unmerge_cells(str(merged_range))
                for r in ws[merged_range.coord]:
                    for c in r:
                        c.value = value

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

# --------------------
# Streamlit 網頁
# --------------------
st.title("班表處理工具 (線上版)")

shift_file = st.file_uploader("上傳班表 Excel 檔案", type=["xlsx","xlsm"])
employee_file = st.file_uploader("上傳員工資料 Excel 檔案", type=["xlsx","xlsm"])

if shift_file and employee_file:
    wb_shift = load_workbook(shift_file)
    excluded = ["彙整結果","班別分析","班別總表"]
    selectable_sheets = [s for s in wb_shift.sheetnames if s not in excluded]

    selected_sheets = st.multiselect("選擇班表工作表", selectable_sheets, default=selectable_sheets)

    wb_employee = load_workbook(employee_file)
    selected_employee_sheet = st.selectbox("選擇員工資料工作表", wb_employee.sheetnames)
    ws_employee = wb_employee[selected_employee_sheet]

    if st.button("開始處理"):
        # 彙整資料
        all_data = []
        for sheet_name in selected_sheets:
            ws = wb_shift[sheet_name]
            unmerge_and_fill(ws)
            clinic_name = str(ws.cell(row=1, column=1).value)[:4]
            max_row, max_col = ws.max_row, ws.max_column
            for r in range(1, max_row+1):
                for c in range(2, max_col+1):
                    cell_value = ws.cell(r, c).value
                    if isinstance(cell_value, datetime):
                        date_val = cell_value
                        i = r+3
                        while i <= max_row:
                            shift_type = str(ws.cell(i, c).value).strip()
                            if isinstance(ws.cell(i, c).value, datetime) or shift_type == "":
                                break
                            if shift_type in ["早","午","晚"]:
                                i +=1
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
                                    i +=1
                                i -=1
                            i +=1
        df_consolidated = pd.DataFrame(all_data, columns=["診所","日期","班別","姓名","A欄資料","U欄資料"])

        # 建立班別分析表
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

        wb_out = Workbook()
        ws_analysis = wb_out.active
        ws_analysis.title = "班別分析"
        headers = ["診所","員工編號","所屬部門","姓名","職稱","日期","班別","E欄資料","班別代碼"]
        ws_analysis.append(headers)

        for key, shift_types in shift_dict.items():
            name, date_str, clinic, e_value = key.split("|")
            shift_type = format_shift_order(shift_types)
            emp_info = emp_dict.get(name, ("","",""))
            emp_id, emp_dept, emp_title = emp_info
            ws_analysis.append([clinic, emp_id, emp_dept, name, emp_title, date_str, shift_type, e_value, get_class_code(emp_title, clinic, shift_type)])

        # 建立班別總表
        all_dates = sorted({row[5] for row in ws_analysis.iter_rows(min_row=2, values_only=True)})
        ws_summary = wb_out.create_sheet("班別總表")
        ws_summary.append(["員工編號","員工姓名"] + all_dates)
        shift_dict_summary = collections.defaultdict(dict)
        for row in ws_analysis.iter_rows(min_row=2, values_only=True):
            emp_id, emp_name, _, _, _, shift_date, _, _, class_code = row[1:]
            emp_key = f"{emp_id}|{emp_name}"
            shift_dict_summary[emp_key][shift_date] = class_code
        for emp_key, date_map in shift_dict_summary.items():
            emp_id, emp_name = emp_key.split("|")
            ws_summary.append([emp_id, emp_name] + [date_map.get(d,"") for d in all_dates])

        # 生成可下載檔案
        output = BytesIO()
        wb_out.save(output)
        output.seek(0)
        st.success("班表處理完成")
        st.download_button("下載結果 Excel", data=output, file_name="output.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
