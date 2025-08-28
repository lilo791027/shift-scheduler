import streamlit as st
import pandas as pd
from openpyxl import load_workbook, Workbook
from datetime import datetime, timedelta
import tempfile
import collections

# --------------------
# 排序班別順序
# --------------------
def format_shift_order(shift_str):
    result = ""
    for s in ["早", "午", "晚"]:
        if s in shift_str:
            result += s
    return result

# --------------------
# 判斷班別代碼
# --------------------
def get_class_code(emp_title, clinic_name, shift_type):
    if not emp_title or pd.isna(emp_title):
        return ""
    # 特殊早班職稱
    if emp_title in ["早班護理師", "早班內視鏡助理", "醫務專員", "兼職早班內視鏡助理"]:
        return "【員工】純早班"

    class_code = ""
    if emp_title == "醫師":
        class_code = "★醫師★"
    elif emp_title in ["櫃臺", "護理師", "兼職護理師", "兼職跟診助理", "副店長", "護士", "藥師"]:
        class_code = "【員工】"
    elif "副店長" in emp_title or emp_title == "採購儲備組長":
        class_code = "【員工】"
    elif "店長" in emp_title:
        class_code = "◇主管◇"

    if shift_type != "早":
        if clinic_name in ["上吉診所", "立吉診所", "上承診所", "立全診所", "立竹診所", "立順診所", "上京診所"]:
            class_code += "板土中京"
        elif clinic_name == "立丞診所":
            class_code += "立丞"

    shift_map = {
        "早": "早班",
        "午": "午班",
        "晚": "晚班",
        "早午": "早午班",
        "午晚": "午晚班",
        "早晚": "早晚班",
        "早午晚": "全天班"
    }
    class_code += shift_map.get(shift_type, shift_type + "班")

    if class_code.endswith("早班早班"):
        class_code = class_code.replace("早班早班", "早班")

    return class_code

# --------------------
# 模組 1：解除合併儲存格並填入原值
# --------------------
def unmerge_and_fill(ws):
    for merged in list(ws.merged_cells.ranges):
        value = ws.cell(merged.min_row, merged.min_col).value
        ws.unmerge_cells(str(merged))
        for row in ws[merged.coord]:
            for cell in row:
                cell.value = value

# --------------------
# 模組 2：整理班別資料（含空白列）
# --------------------
def consolidate_shift_data(ws):
    all_data = []
    clinic_name = str(ws.cell(1, 1).value)[:4]
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
                                ws.cell(i, 21).value if ws.max_column>=21 else ""
                            ])
                            i += 1
                        i -= 1
                    i += 1
    df = pd.DataFrame(all_data, columns=["診所","日期","班別","姓名","A欄資料","U欄資料"])
    return df

# --------------------
# 模組 3：建立班別分析表
# --------------------
def create_shift_analysis(df_shift, df_emp):
    emp_dict = {}
    for _, row in df_emp.iterrows():
        name = str(row['姓名']).strip()
        if name:
            emp_dict[name] = (row.get('員工編號',''), row.get('部門',''), row.get('職稱',''))

    shift_dict = collections.defaultdict(str)
    for _, row in df_shift.iterrows():
        clinic = row['診所']
        date_str = row['日期']
        shift_type = row['班別']
        name = str(row['姓名']).strip()
        if not name or len(name) > 4:
            continue
        key = f"{name}|{date_str}|{clinic}"
        if key not in shift_dict:
            shift_dict[key] = shift_type
        else:
            shift_dict[key] += " " + shift_type

    output_rows = []
    for key, shifts in shift_dict.items():
        name, date_str, clinic = key.split("|")
        shift_type = format_shift_order(shifts)
        emp_info = emp_dict.get(name, ("", "", ""))
        emp_id, emp_dept, emp_title = emp_info
        output_rows.append({
            "診所": clinic,
            "員工編號": emp_id,
            "所屬部門": emp_dept,
            "姓名": name,
            "職稱": emp_title,
            "日期": date_str,
            "班別": shift_type,
            "E欄資料": "",
            "班別代碼": get_class_code(emp_title, clinic, shift_type)
        })

    df_analysis = pd.DataFrame(output_rows)
    return df_analysis

# --------------------
# 模組 4：建立班別總表
# --------------------
def create_shift_summary(df_analysis):
    df_analysis['日期'] = pd.to_datetime(df_analysis['日期'])
    first_date = df_analysis['日期'].min()
    year, month = first_date.year, first_date.month
    days_in_month = (first_date.replace(month=month%12+1, day=1) - timedelta(days=1)).day
    all_dates = [first_date.replace(day=d).strftime("%Y-%m-%d") for d in range(1, days_in_month+1)]

    emp_keys = df_analysis[['員工編號','姓名']].drop_duplicates()
    summary_rows = []
    for _, row in emp_keys.iterrows():
        emp_id, emp_name = row['員工編號'], row['姓名']
        emp_data = df_analysis[df_analysis['員工編號']==emp_id]
        row_dict = {"員工編號": emp_id, "員工姓名": emp_name}
        for d in all_dates:
            val = emp_data.loc[emp_data['日期']==d, '班別代碼']
            row_dict[d] = val.values[0] if not val.empty else ""
        summary_rows.append(row_dict)

    df_summary = pd.DataFrame(summary_rows)
    return df_summary

# --------------------
# Streamlit 主程式
# --------------------
st.title("班表自動化主控程式")

shift_file = st.file_uploader("上傳總表 Excel", type=["xlsx","xlsm"])
employee_file = st.file_uploader("上傳員工資料 Excel", type=["xlsx","xlsm"])

if shift_file and employee_file:
    wb_shift = load_workbook(shift_file)
    ws_shift = wb_shift.active
    df_emp = pd.read_excel(employee_file, sheet_name=0)

    st.info("正在解除合併儲存格...")
    unmerge_and_fill(ws_shift)

    st.info("正在整理班別資料...")
    df_shift = consolidate_shift_data(ws_shift)

    st.info("正在建立班別分析表...")
    df_analysis = create_shift_analysis(df_shift, df_emp)

    st.info("正在建立班別總表...")
    df_summary = create_shift_summary(df_analysis)

    # 儲存 Excel
    tmp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    with pd.ExcelWriter(tmp_file.name, engine='openpyxl') as writer:
        df_shift.to_excel(writer, sheet_name="彙整結果", index=False)
        df_analysis.to_excel(writer, sheet_name="班別分析", index=False)
        df_summary.to_excel(writer, sheet_name="班別總表", index=False)

    tmp_file.close()
    st.success("所有班表任務已完成！")
    with open(tmp_file.name, "rb") as f:
        st.download_button(
            "下載結果 Excel",
            data=f,
            file_name="班表自動化結果.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
