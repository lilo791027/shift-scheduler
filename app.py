import pandas as pd

# =========================
# 讀取 Excel
# =========================
excel_file = "班表.xlsx"
ws_total = pd.read_excel(excel_file, sheet_name="總表")
ws_employee = pd.read_excel(excel_file, sheet_name="員工人事資料明細表")

# =========================
# 模組 1：解除總表合併（openpyxl 可處理，如果需要）
# =========================
# pandas 讀取後一般已經展開合併儲存格，通常不需額外操作

# =========================
# 模組 2：整理班別資料（含空白列）
# =========================
records = []
clinic_name = ws_total.iloc[0, 0][:4]  # 取前 4 個字作診所名

for col in ws_total.columns[1:]:
    for idx, val in ws_total[col].items():
        if pd.to_datetime(val, errors='coerce') is not pd.NaT:
            date_value = val
            i = idx + 3
            while i < len(ws_total):
                shift_type = str(ws_total.iloc[i, ws_total.columns.get_loc(col)]).strip()
                if pd.to_datetime(shift_type, errors='coerce') is not pd.NaT or shift_type == "":
                    break
                if shift_type in ["早", "午", "晚"]:
                    i += 1
                    while i < len(ws_total):
                        cell_val = str(ws_total.iloc[i, 0]).strip()
                        if cell_val in ["早", "午", "晚"] or pd.to_datetime(ws_total.iloc[i, col], errors='coerce') is not pd.NaT:
                            break
                        records.append({
                            "診所": clinic_name,
                            "日期": date_value,
                            "班別": shift_type,
                            "姓名": ws_total.iloc[i, col],
                            "A欄資料": ws_total.iloc[i, 0],
                            "U欄資料": ws_total.iloc[i, 20]  # pandas index 從0開始
                        })
                        i += 1
                    i -= 1
                i += 1

df_shift = pd.DataFrame(records)

# =========================
# 模組 3：建立班別分析表與班別代碼
# =========================
# 建立員工字典
emp_dict = {}
for _, row in ws_employee.iterrows():
    name = str(row[1]).strip()
    if name:
        emp_dict[name] = {
            "empID": str(row[0]),
            "empDept": row[2],
            "empTitle": row[3]
        }

# 合併班別
shift_dict = {}
for _, row in df_shift.iterrows():
    key = f"{row['姓名']}|{row['日期'].strftime('%Y/%m/%d')}|{row['診所']}"
    if key not in shift_dict:
        shift_dict[key] = row['班別']
    else:
        shift_dict[key] += " " + row['班別']

def format_shift_order(s):
    result = ""
    for x in ["早", "午", "晚"]:
        if x in s:
            result += x
    return result

def get_class_code(empTitle, clinicName, shiftType):
    if not empTitle or pd.isna(empTitle):
        return ""
    if empTitle in ["早班護理師", "早班內視鏡助理", "醫務專員", "兼職早班內視鏡助理"]:
        return "【員工】純早班"
    
    classCode = ""
    if empTitle == "醫師":
        classCode = "★醫師★"
    elif empTitle in ["櫃臺", "護理師", "兼職護理師", "兼職跟診助理", "副店長", "護士", "藥師"] or "副店長" in empTitle:
        classCode = "【員工】"
    elif "店長" in empTitle or "採購儲備組長" in empTitle:
        classCode = "◇主管◇"

    if shiftType != "早":
        if clinicName in ["上吉診所", "立吉診所", "上承診所", "立全診所", "立竹診所", "立順診所", "上京診所"]:
            classCode += "板土中京"
        elif clinicName == "立丞診所":
            classCode += "立丞"

    mapping = {
        "早": "早班",
        "午晚": "午晚班",
        "早午晚": "全天班",
        "早晚": "早晚班",
        "午": "午班",
        "晚": "晚班",
        "早午": "早午班"
    }
    classCode += mapping.get(shiftType, "")
    return classCode.replace("早班早班", "早班")

# 輸出班別分析表
records_analysis = []
for key, shiftType in shift_dict.items():
    name, dateValue, clinicName = key.split("|")
    shiftType = format_shift_order(shiftType)
    emp_info = emp_dict.get(name, {"empID": "", "empDept": "", "empTitle": ""})
    records_analysis.append({
        "診所": clinicName,
        "員工編號": emp_info["empID"],
        "所屬部門": emp_info["empDept"],
        "姓名": name,
        "職稱": emp_info["empTitle"],
        "日期": dateValue,
        "班別": shiftType,
        "E欄資料": "",
        "班別代碼": get_class_code(emp_info["empTitle"], clinicName, shiftType)
    })

df_analysis = pd.DataFrame(records_analysis)

# =========================
# 模組 4：建立班別總表
# =========================
df_analysis['日期'] = pd.to_datetime(df_analysis['日期'])
df_summary = df_analysis.pivot_table(
    index=['員工編號', '姓名'],
    columns='日期',
    values='班別代碼',
    aggfunc='first'
).reset_index()

# =========================
# 寫入 Excel
# =========================
with pd.ExcelWriter("班表結果.xlsx", engine="openpyxl") as writer:
    df_shift.to_excel(writer, sheet_name="彙整結果", index=False)
    df_analysis.to_excel(writer, sheet_name="班別分析", index=False)
    df_summary.to_excel(writer, sheet_name="班別總表", index=False)

print("班表處理完成！")
