import streamlit as st
import pandas as pd
import io

# ------------------- 函數區 -------------------
def unmerge_and_fill(df):
    return df.ffill(axis=0).ffill(axis=1)

def summarize_schedule(df):
    result = []
    clinic_name = str(df.iloc[0,0])[:4] if not df.empty else ""
    for c in df.columns[1:]:
        for r in range(df.shape[0]):
            try:
                if pd.to_datetime(df.iloc[r,c], errors='coerce') is not pd.NaT:
                    date_value = pd.to_datetime(df.iloc[r,c])
                    i = r + 3
                    while i < df.shape[0]:
                        shift_type = str(df.iloc[i,c]).strip()
                        if pd.to_datetime(df.iloc[i,c], errors='coerce') is not pd.NaT or shift_type == "":
                            break
                        if shift_type in ["早", "午", "晚"]:
                            i += 1
                            while i < df.shape[0]:
                                if pd.to_datetime(df.iloc[i,c], errors='coerce') is not pd.NaT:
                                    break
                                cell_value = str(df.iloc[i,c]).strip()
                                if cell_value in ["早", "午", "晚"]:
                                    break
                                result.append([
                                    clinic_name,
                                    date_value.strftime("%Y/%m/%d"),
                                    shift_type,
                                    cell_value,
                                    df.iloc[i,0],
                                    df.iloc[i,20] if df.shape[1] > 20 else ""
                                ])
                                i += 1
                            i -= 1
                        i += 1
            except:
                continue
    return pd.DataFrame(result, columns=["診所","日期","班別","員工姓名","A欄資料","U欄資料"])

def get_class_code(empTitle, clinicName, shiftType):
    if not empTitle:
        return ""
    class_code = ""
    if empTitle in ["早班護理師", "早班內視鏡助理", "醫務專員", "兼職早班內視鏡助理"]:
        return "【員工】純早班"
    if empTitle == "醫師":
        class_code = "★醫師★"
    elif empTitle in ["櫃臺","護理師","兼職護理師","兼職跟診助理","副店長"] or "副店長" in empTitle:
        class_code = "【員工】"
    elif "店長" in empTitle or "護士" in empTitle:
        class_code = "◇主管◇"
    if shiftType != "早":
        if clinicName in ["上吉診所","立吉診所","上承診所","立全診所","立竹診所","立順診所","上京診所"]:
            class_code += "板土中京"
        elif clinicName == "立丞診所":
            class_code += "立丞"
    mapping = {"早":"早班","午晚":"午晚班","早午晚":"全天班","早晚":"早晚班","午":"午班","晚":"晚班","早午":"早午班"}
    class_code += mapping.get(shiftType, shiftType)
    class_code = class_code.replace("早班早班","早班")
    return class_code

def build_shift_analysis(summarized_df, employee_df):
    summarized_df['員工姓名'] = summarized_df['員工姓名'].astype(str)
    employee_df.columns = employee_df.columns.str.strip()
    emp_dict = {row['員工姓名']: (str(row['員工編號']), row['所屬部門'], row['職稱'])
                for idx,row in employee_df.iterrows()}
    shift_dict = {}
    for _, row in summarized_df.iterrows():
        key = f"{row['員工姓名']}|{row['日期']}|{row['診所']}|{row['A欄資料']}"
        shift_dict[key] = shift_dict.get(key, "") + row['班別']

    analysis_rows = []
    for key, shift_str in shift_dict.items():
        name, date_value, clinic_name, e_value = key.split("|")
        empID, empDept, empTitle = emp_dict.get(name, ("", "", ""))
        shift_order = "".join([s for s in ["早","午","晚"] if s in shift_str])
        class_code = get_class_code(empTitle, clinic_name, shift_order)
        analysis_rows.append([clinic_name, empID, empDept, name, empTitle, date_value, shift_order, e_value, class_code])

    return pd.DataFrame(analysis_rows, columns=["診所","員工編號","所屬部門","員工姓名","職稱","日期","班別","E欄資料","班別代碼"])

def build_shift_summary(analysis_df):
    all_dates = sorted(analysis_df['日期'].unique())
    summary_dict = {}
    for _, row in analysis_df.iterrows():
        emp_key = (row['員工編號'], row['員工姓名'])
        summary_dict.setdefault(emp_key,{})[row['日期']] = row['班別代碼']

    summary_rows = []
    for (empID, empName), date_dict in summary_dict.items():
        row = [empID, empName] + [date_dict.get(d,"") for d in all_dates]
        summary_rows.append(row)
    return pd.DataFrame(summary_rows, columns=["員工編號","員工姓名"] + list(all_dates))

# ------------------- Streamlit 網頁 -------------------
st.title("線上排班系統（偵錯版）")

schedule_file = st.file_uploader("班表 Excel", type=["xlsx"])
employee_file = st.file_uploader("員工資料 Excel", type=["xlsx"])

if schedule_file and employee_file:
    df_schedule = pd.read_excel(schedule_file)
    df_employee = pd.read_excel(employee_file)

    st.subheader("1️⃣ 原始班表資料")
    st.dataframe(df_schedule)

    st.subheader("2️⃣ 原始員工資料")
    st.dataframe(df_employee)

    df_schedule = unmerge_and_fill(df_schedule)
    df_summary = summarize_schedule(df_schedule)

    st.subheader("3️⃣ 彙整結果（df_summary）")
    st.dataframe(df_summary)

    df_analysis = build_shift_analysis(df_summary, df_employee)

    st.subheader("4️⃣ 班別分析（df_analysis）")
    st.dataframe(df_analysis)

    df_final = build_shift_summary(df_analysis)

    st.subheader("5️⃣ 班別總表（df_final）")
    st.dataframe(df_final)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_summary.to_excel(writer, sheet_name="彙整結果", index=False)
        df_analysis.to_excel(writer, sheet_name="班別分析", index=False)
        df_final.to_excel(writer, sheet_name="班別總表", index=False)
    output.seek(0)

    st.download_button(
        label="📥 下載排班結果 Excel",
        data=output,
        file_name="排班結果.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
