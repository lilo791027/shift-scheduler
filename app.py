import os
import platform
import subprocess
from datetime import datetime
import pandas as pd
import collections
from openpyxl import load_workbook
from tkinter import Tk, filedialog, Listbox, MULTIPLE, Button, Toplevel, END

# --------------------
# 模組 1: 解合併並填入原值
# --------------------
def unmerge_and_fill(ws):
    try:
        for merged_range in list(ws.merged_cells.ranges):
            value = ws[merged_range.min_row][merged_range.min_col].value
            ws.unmerge_cells(str(merged_range))
            for row in ws[merged_range.coord]:
                for cell in row:
                    cell.value = value
    except Exception as e:
        print("解合併錯誤:", e)

# --------------------
# GUI 多選工作表
# --------------------
def select_sheets_gui(sheet_names, title="選擇工作表"):
    selected = []

    def on_select():
        for idx in listbox.curselection():
            selected.append(sheet_names[idx])
        top.destroy()

    top = Toplevel()
    top.title(title)
    top.geometry("300x400")
    listbox = Listbox(top, selectmode=MULTIPLE)
    listbox.pack(expand=True, fill='both')

    for s in sheet_names:
        listbox.insert(END, s)

    btn = Button(top, text="確認", command=on_select)
    btn.pack()
    top.wait_window()
    return selected

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
                                val = str(ws.cell(i, c).value).strip() if ws.cell(i, c).value else ""
                                if isinstance(ws.cell(i, c).value, datetime) or val in ["早","午","晚"]:
                                    break
                                all_data.append([
                                    clinic_name,
                                    date_val.strftime("%Y/%m/%d"),
                                    shift_type,
                                    val,
                                    ws.cell(i, 1).value if ws.cell(i, 1).value else "",
                                    ws.cell(i, 21).value if ws.cell(i, 21).value else ""
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

def create_shift_analysis(wb_shift, df_consolidated, ws_employee):
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

    if "班別分析" in wb_shift.sheetnames:
        ws_target = wb_shift["班別分析"]
        ws_target.delete_rows(1, ws_target.max_row)
    else:
        ws_target = wb_shift.create_sheet("班別分析")

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
def create_shift_summary(wb_shift):
    ws_analysis = wb_shift["班別分析"]
    all_dates = [datetime(2025,8,i).strftime("%Y-%m-%d") for i in range(1,32)]
    shift_dict = collections.defaultdict(dict)

    for row in ws_analysis.iter_rows(min_row=2, values_only=True):
        row = list(row) + [""]*9
        emp_id = row[1] or ""
        emp_name = row[3] or ""
        shift_date = row[5] or ""
        class_code = row[8] or ""
        if not emp_id or not emp_name or not shift_date:
            continue
        emp_key = f"{emp_id}|{emp_name}"
        shift_dict[emp_key][shift_date] = class_code

    if "班別總表" in wb_shift.sheetnames:
        ws_target = wb_shift["班別總表"]
        ws_target.delete_rows(1, ws_target.max_row)
    else:
        ws_target = wb_shift.create_sheet("班別總表")

    ws_target.append(["員工編號","員工姓名"] + all_dates)
    for emp_key, date_map in shift_dict.items():
        emp_id, emp_name = emp_key.split("|")
        row = [emp_id, emp_name] + [date_map.get(d,"") for d in all_dates]
        ws_target.append(row)

# --------------------
# 主程式
# --------------------
def run_all():
    Tk().withdraw()
    # 選班表檔案
    shift_file = filedialog.askopenfilename(title="選擇班表 Excel 檔案", filetypes=[("Excel Files","*.xlsx *.xlsm")])
    if not shift_file:
        print("未選擇班表檔案，程式結束。")
        return
    wb_shift = load_workbook(shift_file)

    excluded = ["彙整結果","班別分析","班別總表"]
    selectable_sheets = [s for s in wb_shift.sheetnames if s not in excluded]
    selected_sheets = select_sheets_gui(selectable_sheets, title="選擇班表工作表")
    if not selected_sheets:
        print("未選擇工作表，程式結束。")
        return

    # 選員工資料檔案
    employee_file = filedialog.askopenfilename(title="選擇員工資料 Excel 檔案", filetypes=[("Excel Files","*.xlsx *.xlsm")])
    if not employee_file:
        print("未選擇員工資料檔案，程式結束。")
        return
    wb_employee = load_workbook(employee_file)
    employee_sheet_name = select_sheets_gui(wb_employee.sheetnames, title="選擇員工資料工作表")[0]
    ws_employee = wb_employee[employee_sheet_name]

    df_consolidated = consolidate_selected_sheets(wb_shift, selected_sheets)

    # 彙整結果工作表
    if "彙整結果" in wb_shift.sheetnames:
        ws_out = wb_shift["彙整結果"]
        ws_out.delete_rows(1, ws_out.max_row)
    else:
        ws_out = wb_shift.create_sheet("彙整結果")

    for r_idx, row in df_consolidated.iterrows():
        for c_idx, val in enumerate(row, 1):
            ws_out.cell(row=r_idx+2, column=c_idx, value=val)

    # 建立分析表與總表
    create_shift_analysis(wb_shift, df_consolidated, ws_employee)
    create_shift_summary(wb_shift)

    # 存檔到桌面
    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    output_path = os.path.join(desktop, "output.xlsx")
    wb_shift.save(output_path)
    print(f"完成！結果存為: {output_path}")

    # 自動打開 Excel
    try:
        if platform.system() == "Windows":
            os.startfile(output_path)
        elif platform.system() == "Darwin":
            subprocess.call(["open", output_path])
        else:
            subprocess.call(["xdg-open", output_path])
    except:
        print("無法自動開啟 Excel，請手動打開 output.xlsx 查看結果")

# ===== 執行 =====
run_all()
