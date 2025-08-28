import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from io import BytesIO
from datetime import datetime

# --------------------
# 生成班別分析表
# --------------------
def create_shift_analysis(df_shift: pd.DataFrame, df_emp: pd.DataFrame) -> pd.DataFrame:
    df_shift = df_shift.copy()
    df_emp = df_emp.copy()
    df_shift.columns = [str(c).strip() for c in df_shift.columns]
    df_emp.columns = [str(c).strip() for c in df_emp.columns]

    # 自動匹配欄位
    shift_col_map = {}
    for col in df_shift.columns:
        if "姓名" in col:
            shift_col_map["name"] = col
        elif "班別" in col:
            shift_col_map["shift"] = col
        elif "診所" in col:
            shift_col_map["clinic"] = col
        elif "日期" in col:
            shift_col_map["date"] = col

    emp_col_map = {}
    for col in df_emp.columns:
        if "姓名" in col:
            emp_col_map["name"] = col
        elif "員工編號" in col or "ID" in col:
            emp_col_map["id"] = col
        elif "職稱" in col:
            emp_col_map["title"] = col

    emp_dict = {}
    for _, row in df_emp.iterrows():
        name = str(row.get(emp_col_map.get("name"), "")).strip()
        if name:
            emp_dict[name] = (str(row.get(emp_col_map.get("id"), "")).strip(), str(row.get(emp_col_map.get("title"), "")).strip())

    shift_dict = {}
    for _, row in df_shift.iterrows():
        name = str(row.get(shift_col_map.get("name"), "")).strip()
        clinic = str(row.get(shift_col_map.get("clinic"), "")).strip()
        date_val = row.get(shift_col_map.get("date"), "")
        shift_type = str(row.get(shift_col_map.get("shift"), "")).strip()

        if not name:
            continue

        key = f"{name}|{date_val}|{clinic}"
        if key not in shift_dict:
            shift_dict[key] = shift_type
        else:
            shift_dict[key] += " " + shift_type

    data_out = []
    for key, shifts in shift_dict.items():
        name, date_val, clinic = key.split("|")
        emp_id, emp_title = emp_dict.get(name, ("", ""))
        shift_type = "".join([s for s in ["早", "午", "晚"] if s in shifts])
        data_out.append([emp_id, name, date_val, shift_type])

    df_analysis = pd.DataFrame(data_out, columns=["員工編號", "員工姓名", "日期", "班別"])
    return df_analysis

# --------------------
# 生成班別總表
# --------------------
def create_shift_summary(df_analysis: pd.DataFrame) -> pd.DataFrame:
    if df_analysis.empty:
        return pd.DataFrame()
    df_analysis["日期"] = pd.to_datetime(df_analysis["日期"])
    all_dates = pd.date_range(df_analysis["日期"].min(), df_analysis["日期"].max()).strftime("%Y-%m-%d").tolist()

    summary_dict = {}
    for _, row in df_analysis.iterrows():
        emp_id = str(row["員工編號"])
        emp_name = row["員工姓名"]
        shift_date = row["日期"].strftime("%Y-%m-%d")
        class_code = row["班別"]

        key = (emp_id, emp_name)
        if key not in summary_dict:
            summary_dict[key] = {}
        summary_dict[key][shift_date] = class_code

    data_out = []
    for (emp_id, emp_name), shifts in summary_dict.items():
        row = [emp_id, emp_name] + [shifts.get(d, "") for d in all_dates]
        data_out.append(row)

    columns = ["員工編號", "員工姓名"] + all_dates
    df_summary = pd.DataFrame(data_out, columns=columns)
    return df_summary

# --------------------
# Streamlit 主程式
# --------------------
st.title("上吉醫療-班表轉換鋒形格式")

shift_file = st.file_uploader("上傳班表 Excel 檔案", type=["xlsx", "xlsm"])
employee_file = st.file_uploader("上傳員工資料 Excel 檔案", type=["xlsx", "xlsm"])

if shift_file and employee_file:
    df_shift = pd.read_excel(shift_file)
    df_emp = pd.read_excel(employee_file)

    if st.button("生成班別總表"):
        df_analysis = create_shift_analysis(df_shift, df_emp)
        df_summary = create_shift_summary(df_analysis)

        st.success("班別總表生成完成！")
        st.dataframe(df_summary)

        with BytesIO() as output:
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                df_summary.to_excel(writer, sheet_name="班別總表", index=False)
            st.download_button("下載班別總表 Excel", data=output.getvalue(), file_name="班別總表.xlsx")
