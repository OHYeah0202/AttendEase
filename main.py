import math
import shutil
import logging
import traceback
from datetime import datetime
from typing import Dict, Any

import pandas as pd
import tkinter as tk
from tkinter import messagebox

from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import PatternFill, Font
from tqdm import tqdm
from datetime import time
import sys
import os

if getattr(sys, 'frozen', False):
    # 如果是打包成exe後執行
    base_path = os.path.dirname(sys.executable)
else:
    # 如果是Python開發模式
    base_path = os.path.dirname(os.path.abspath(__file__))

# 建立 log 資料夾與設定 logging
log_dir = os.path.join(base_path, 'log')
os.makedirs(log_dir, exist_ok=True)

# log 檔名
log_file = os.path.join(log_dir, f"{datetime.now().strftime('%Y%m%d%H%M%S')}_attendease.log")

# 建立 Logger
logger = logging.getLogger()
logger.setLevel(logging.INFO)

# 格式設定
formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')

# FileHandler
file_handler = logging.FileHandler(log_file, encoding='utf-8')
file_handler.setFormatter(formatter)

# ConsoleHandler
console_handler = logging.StreamHandler()
console_handler.setFormatter(formatter)

# 避免重複加入 Handler（pyinstaller EXE 或重啟時）
if not logger.hasHandlers():
    logger.addHandler(file_handler)
    logger.addHandler(console_handler)

# 基礎設定
SHIFT_RULES = {
    'A1': { # DL 一般員工
        'start': time(7, 40),
        'end': time(18, 40),
        'break': 0.83 + 0.67 + 0.33, # 午餐 + 晚餐 + 小休
        'friday_end': time(18, 40) # 跟平日一樣
    },
    'A2': { # DL 巫裔男性，週五延後下班
        'start': time(7, 40),
        'end': time(18, 40),
        'break': 0.83 + 0.67 + 0.33, # 午餐 + 晚餐 + 小休
        'friday_end': time(19, 40) # 周五特別處理
    },
    'B1': { # IDL 員工
        'start': time(8, 0),
        'end': time(18, 0),
        'break': 1.0,
        'friday_end': None  # 周五不需要特別處理
    }
}

DL_WORK_START = pd.to_datetime('07:50:00').time()  # DL 上班時間
DL_WORK_END = pd.to_datetime('18:40:00').time()  # DL 下班時間
DL_WORK_END_FRIDAY = pd.to_datetime('19:40:00').time()  # DL(男) 周五下班時間
IDL_WORK_START = pd.to_datetime('08:00:00').time()  # IDL 上班時間
IDL_WORK_END = pd.to_datetime('18:00:00').time()  # IDL 下班時間

DL_LUNCH_BREAK = 0.83  # DL 50 分鐘午餐
DL_DINNER_BREAK = 0.67  # DL 40 分鐘晚餐
DL_SMALL_BREAK = 0.33  # DL 20 分鐘小休息
IDL_LUNCH_BREAK = 1  # IDL 1 小時午餐


def load_excel_file(filename):
    path = os.path.join(base_path, 'data', filename)
    return pd.read_excel(path, sheet_name=None)


# 計算遲到
def calc_late(row):
    if pd.isna(row['Clock-in']):
        return "-"
    if row['Clock-in'] > IDL_WORK_START and row['DAY TYPE'] in ['WORK', 'OT']:
        late_min = (datetime.combine(datetime.today(), row['Clock-in']) - datetime.combine(datetime.today(),
                                                                                           IDL_WORK_START)).seconds // 60
        return late_min
    else:
        return "-"


# 計算早退
def calc_early(row):
    if pd.isna(row['Clock-out']):
        return "-"
    if row['Clock-out'] < IDL_WORK_END and row['DAY TYPE'] in ['WORK', 'OT']:
        return (datetime.combine(datetime.today(), IDL_WORK_END) - datetime.combine(datetime.today(),
                                                                                    row['Clock-out'])).seconds // 60
    else:
        return "-"


# 計算工時
def calc_work_hours(row):
    if pd.isna(row['Clock-in']) or pd.isna(row['Clock-out']):
        return 0.0

    shift_code = map_shift(row)
    shift_info = SHIFT_RULES.get(shift_code)

    if not shift_info:
        return 0

    is_friday = row.get('Day', '').strip().lower() == 'fri.'

    scheduled_start = shift_info['start']
    total_break = shift_info['break']

    actual_start = max(row['Clock-in'], scheduled_start)
    actual_end = row['Clock-out']

    work_start_dt = datetime.combine(datetime.today(), actual_start)
    work_end_dt = datetime.combine(datetime.today(), actual_end)
    work_hours = (work_end_dt - work_start_dt).total_seconds() / 3600

    if is_friday and shift_info['friday_end'] == time(19, 40):
        net_hours = max(0.0, work_hours - total_break - 1.0) #
        return round(net_hours, 2)

    else:
        net_hours = max(0.0, work_hours - total_break)
        return round(net_hours, 2)


# 計算加班
def calc_ot(row):
    shift_code = map_shift(row)

    if shift_code is None:
        return 0

    shift_code = shift_code.strip().upper()
    work_hours = row['WORK']
    day = row['Day'].strip().lower()

    if shift_code in ['A1', 'A2']:
        if row['DAY TYPE'] == "PH":
            ot_units = math.floor(work_hours * 60 / 30) * 0.5

            if ot_units > 8:
                monthly_counters['OT2.0'] += 8
                monthly_counters['OT3.0'] += ot_units - 8
            else:
                monthly_counters['OT2.0'] += ot_units

            return ot_units

        elif day == "sun.":
            ot_units = math.floor(work_hours * 60 / 30) * 0.5
            monthly_counters['OT2.0'] += ot_units
            return ot_units

        elif day == 'sat.':
            ot_units = math.floor(work_hours * 60 / 30) * 0.5
            monthly_counters['OT1.5'] += ot_units
            return ot_units

        else:
            ot_hours = max(0, work_hours - 9)  # 超出標準9小時才算加班
            ot_units = math.floor(ot_hours * 60 / 30) * 0.5  # 換成小數點
            monthly_counters['OT1.5'] += ot_units
            return ot_units

    else:
        return 0


# 合并請假資料（簡單版：直接貼在對應日）
def map_leave(row):
    record = leave_df[(leave_df['Employee ID'] == row['Employee ID']) &
                      (leave_df['Start Date'] == row['Date'])]

    if not record.empty:
        return record.iloc[0]['Leave Type']
    else:
        if row['DAY TYPE'] == "OT" and (pd.isna(row['Clock-in']) or pd.isna(row['Clock-out']) or row['Clock-out'] < pd.to_datetime('18:00:00').time()):
            monthly_counters['CANNOT_OT'] += 1
            return "Cannot OT"

        return "-"


# 合并請假天數
def map_lve_days(row):
    record = leave_df[(leave_df['Employee ID'] == row['Employee ID']) &
                      (leave_df['Start Date'] == row['Date'])]
    if not record.empty:
        return record.iloc[0]['Days']
    else:
        return 0


# 判斷是否有午餐津貼
def map_meal(row):
    if pd.isna(row['Clock-in']) or pd.isna(row['Clock-out']):
        return 0

    record = meal_df[(meal_df['Date'] == row['Date']) &
                     (meal_df['Employee ID'] == row['Employee ID'])]

    if row['DAY TYPE'] == 'WORK' and record.empty and row['LATE_MIN'] == '-' and row['Clock-out'] >= pd.to_datetime('18:00:00').time():
        monthly_counters['MEAL'] += 3
        return 3
    return 0


# 判斷工作日、假日、休假日、加班日
def auto_day_type(row):
    shift_code = map_shift(row)
    day = row['Day']

    if shift_code is None:
        return 'OFF'

    shift_code = shift_code.strip().upper()
    personal_record = holiday_df[
        (holiday_df['Employee ID'] == row['Employee ID']) &
        (holiday_df['Date'] == row['Date'])
    ]

    if not personal_record.empty:
        return personal_record.iloc[0]['Festival Name']

    general_record = holiday_df[
        (holiday_df['Employee ID'].isna()) &
        (holiday_df['Date'] == row['Date'])
    ]

    if not general_record.empty:
        festival_name = general_record.iloc[0]['Festival Name']

        if festival_name == 'OFF' and shift_code in ['A1', 'A2']:
            return 'OFF'

        return 'PH'

    if day == "Sun.":
        return 'REST'

    if day == "Sat.":
        return 'OFF' if shift_code == 'B1' else 'OT'  # B1 = IDL => 週六排休；B2 = DL => 週六加班

    return 'WORK'


# 判斷是否有手動輸入的加班資料
def map_manual_ot(row):
    record = manualot_df[(manualot_df['Employee ID'] == row['Employee ID']) &
                         (manualot_df['Date'] == row['Date'])]

    if not record.empty:
        return record.iloc[0]['OT Minutes']
    else:
        return '-'


# 判斷班別
def map_shift(row):
    emp_info = employee_df.loc[employee_df['Employee ID'] == row['Employee ID']]

    if emp_info.empty:
        return None

    return emp_info.iloc[0]['Shift']


# 取得欄位 index
def get_col_index(name):
    return header.index(name) if name in header else None


# 找尋 Excel 第一個空白列
def find_first_empty_row(sheet, id_column_index):
    for row_num in range(2, sheet.max_row + 1):
        cell = sheet.cell(row=row_num, column=id_column_index)
        if not cell.value:
            return row_num
    return sheet.max_row + 1


for step in tqdm(range(2), desc="🚀 In Progress..."):
    if step == 0:
        logging.info("📁 Reading masterdata.xlsx...")
        # 讀取資料
        xls = load_excel_file('masterdata.xlsx')

        try:
            employee_df = xls.get('Employee')
            attendance_df = xls.get('Attendance')
            leave_df = xls.get('Leave')
            holiday_df = xls.get('Holiday')
            meal_df = xls.get('Meal')
            manualot_df = xls.get('Manual OT')
        except Exception as e:
            logging.error("Fail to read worksheet！" + str(e))
            messagebox.showerror("Error", f"An error occurred：Fail to read worksheet.\nPlease check the log file！")

    elif step == 1:
        logging.info("📊️ Data Processing...")
        # 設定檔案路徑
        OUTPUT_FOLDER = 'output'
        TEMPLATE_FOLDER = 'template'

        # 檢查 output 和 template 資料夾
        os.makedirs(OUTPUT_FOLDER, exist_ok=True)
        os.makedirs(TEMPLATE_FOLDER, exist_ok=True)

        current_month = datetime.now().strftime('%B')

        EMPLOYEE_TEMPLATE_PATH = os.path.join(TEMPLATE_FOLDER, "employee_report_template.xlsx")
        MASTER_TEMPLATE_PATH = os.path.join(TEMPLATE_FOLDER, "master_report_template.xlsx")
        # EMPLOYEE_OUTPUT_PATH = os.path.join(OUTPUT_FOLDER, f"{current_month}_Employees_Report.xlsx")
        MASTER_OUTPUT_PATH = os.path.join(OUTPUT_FOLDER, "Master_Report.xlsx")

        # 建立 output 檔案 (不動到 template，複製一份)
        try:
            # shutil.copy(EMPLOYEE_TEMPLATE_PATH, EMPLOYEE_OUTPUT_PATH)
            shutil.copy(MASTER_TEMPLATE_PATH, MASTER_OUTPUT_PATH)
            # wb_employee = load_workbook(EMPLOYEE_OUTPUT_PATH)
            wb_master = load_workbook(MASTER_OUTPUT_PATH)
        except Exception as e:
            logging.error("❌ An error occurred while duplicate report template！")
            logging.error(traceback.format_exc())
            messagebox.showerror("Error", f"An error occurred：{str(e)}\nPlease check the log file！")

        # 顏色設定
        yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        pink_fill = PatternFill(start_color="FFC0CB", end_color="FFC0CB", fill_type="solid")

        # 格式轉換
        try:
            attendance_df['Clock-in'] = pd.to_datetime(attendance_df['Clock-in'], errors='coerce').dt.time
            attendance_df['Clock-out'] = pd.to_datetime(attendance_df['Clock-out'], errors='coerce').dt.time

            leave_df['Start Date'] = pd.to_datetime(leave_df['Start Date'])
            leave_df['End Date'] = pd.to_datetime(leave_df['End Date'])
        except Exception as e:
            logging.error("❌ An error occurred while executing the program！")
            logging.error(traceback.format_exc())
            messagebox.showerror("Error", f"An error occurred：{str(e)}\nPlease check the log file！")

        try:
            dept_grouped = attendance_df.groupby('Company / Department')

            for dept_name, dept_group in dept_grouped:
                logging.info(f"🏢 Processing Department: {dept_name}")

                safe_dept_name = dept_name.replace('/', '_').replace('\\', '_')
                DEPARTMENT_OUTPUT_PATH = os.path.join(OUTPUT_FOLDER, f"{current_month}_{safe_dept_name}_Report.xlsx")

                shutil.copy(EMPLOYEE_TEMPLATE_PATH, DEPARTMENT_OUTPUT_PATH)
                wb_dept = load_workbook(DEPARTMENT_OUTPUT_PATH)

                emp_grouped = dept_group.groupby('Employee ID')

                for emp_id, group in emp_grouped:
                    group = group.sort_values(by='Date').reset_index(drop=True)

                    # 每月統計資料
                    monthly_counters = {
                        'LATE_IN': 0,
                        'EARLY_OUT': 0,
                        'FORGOT_CLOCKING': 0,
                        'ABSENT': 0,
                        'WORK_HOURS': 0,
                        'OT_HOURS': 0,
                        'LVE DAYS': 0,
                        'MEAL': 0,
                        'OT1.5': 0,
                        'OT2.0': 0,
                        'OT3.0': 0,
                        'CANNOT_OT': 0,
                        'MANUAL_OT': 0,
                        'FINAL_OT1.5': 0
                    }

                    # 整理欄位和新增欄位
                    group['DAY TYPE'] = group.apply(auto_day_type, axis=1)
                    group['WORK'] = group.apply(calc_work_hours, axis=1)
                    group['OT'] = group.apply(calc_ot, axis=1)
                    group['LATE_MIN'] = group.apply(calc_late, axis=1)
                    group['EARLY_MIN'] = group.apply(calc_early, axis=1)
                    group['LEAVE'] = group.apply(map_leave, axis=1)
                    group['LVE DAYS'] = group.apply(map_lve_days, axis=1)
                    # 忘記打卡
                    group['FORGOT_CLOCKING'] = group.apply(
                        lambda row: 1 if row['DAY TYPE'] == "WORK" and (
                                pd.isna(row['Clock-in']) or pd.isna(row['Clock-out'])) and not (
                                    pd.isna(row['Clock-in']) and pd.isna(row['Clock-out'])) and row['Day'] not in [
                                             'Sat.', 'Sun.'] else 0, axis=1)
                    # 計算缺勤(ABSENT)
                    group['ABSENT'] = group.apply(
                        lambda row: 1 if (
                                row['DAY TYPE'] == "WORK" and (
                                    pd.isna(row['Clock-in']) and pd.isna(row['Clock-out'])) and
                                row['LEAVE'] == '-') else 0, axis=1)
                    # 判斷班別
                    group['SHIFT'] = group.apply(map_shift, axis=1)

                    if group.iloc[0]['SHIFT'] not in ['A1', 'A2', 'B1']:
                        logging.warning(f"️⚠️ Cannot determine shift code for Employee ID: {emp_id}")

                    group['MEAL'] = group.apply(map_meal, axis=1)
                    group['MANUAL OT'] = group.apply(map_manual_ot, axis=1)

                    # 插入每周統計小結
                    final_rows = []

                    weekly_counters = {
                        'LATE_IN': 0,
                        'EARLY_OUT': 0,
                        'FORGOT_CLOCKING': 0,
                        'ABSENT': 0,
                        'WORK_HOURS': 0,
                        'OT_HOURS': 0,
                        'LVE DAYS': 0,
                        'MEAL': 0,
                        'MANUAL OT': 0
                    }

                    selected_columns = [
                        "Date", "Day", "DAY TYPE", "Clock-in", "Clock-out", "SHIFT",
                        "LATE_MIN", "EARLY_MIN", "FORGOT_CLOCKING", "ABSENT", "WORK",
                        "LEAVE", "LVE DAYS", "MEAL", "OT", 'MANUAL OT'
                    ]

                    for idx, row in group.iterrows():
                        filtered_row = {col: row[col] for col in selected_columns if col in row}
                        filtered_row['Date'] = row['Date'].strftime('%Y-%m-%d')
                        final_rows.append(filtered_row)

                        # 每月統計
                        monthly_counters['LATE_IN'] += 0 if row['LATE_MIN'] == '-' else row['LATE_MIN']
                        monthly_counters['EARLY_OUT'] += 0 if row['EARLY_MIN'] == '-' else row['EARLY_MIN']
                        monthly_counters['FORGOT_CLOCKING'] += row['FORGOT_CLOCKING']
                        monthly_counters['ABSENT'] += row['ABSENT']
                        monthly_counters['WORK_HOURS'] += row['WORK']
                        monthly_counters['OT_HOURS'] += row['OT']
                        monthly_counters['LVE DAYS'] += row['LVE DAYS']
                        monthly_counters['MANUAL_OT'] += 0 if row['MANUAL OT'] == '-' else row['MANUAL OT']

                        # 每周統計
                        weekly_counters['LATE_IN'] += 0 if row['LATE_MIN'] == '-' else row['LATE_MIN']
                        weekly_counters['EARLY_OUT'] += 0 if row['EARLY_MIN'] == '-' else row['EARLY_MIN']
                        weekly_counters['FORGOT_CLOCKING'] += row['FORGOT_CLOCKING']
                        weekly_counters['ABSENT'] += row['ABSENT']
                        weekly_counters['WORK_HOURS'] += row['WORK']
                        weekly_counters['OT_HOURS'] += row['OT']
                        weekly_counters['LVE DAYS'] += row['LVE DAYS']
                        weekly_counters['MEAL'] += row['MEAL']
                        weekly_counters['MANUAL OT'] += 0 if row['MANUAL OT'] == '-' else row['MANUAL OT']

                        if row['Day'] == "Sun." or idx == len(group) - 1:
                            week_summary_row = {
                                'Date': '',
                                'Day': f"Summary up to {row['Date'].strftime('%Y-%m-%d')}",
                                'DAY TYPE': '',
                                'Clock-in': '',
                                'Clock-out': '',
                                'SHIFT': '',
                                'LATE_MIN': weekly_counters['LATE_IN'],
                                'EARLY_MIN': weekly_counters['EARLY_OUT'],
                                'FORGOT_CLOCKING': weekly_counters['FORGOT_CLOCKING'],
                                'ABSENT': weekly_counters['ABSENT'],
                                'WORK': weekly_counters['WORK_HOURS'],
                                'LEAVE': '',
                                'LVE DAYS': weekly_counters['LVE DAYS'],
                                'MEAL': weekly_counters['MEAL'],
                                'OT': weekly_counters['OT_HOURS'],
                                'MANUAL OT': weekly_counters['MANUAL OT']
                            }
                            final_rows.append(pd.Series(week_summary_row))

                            # 重置每周統計
                            weekly_counters = {
                                'LATE_IN': 0,
                                'EARLY_OUT': 0,
                                'FORGOT_CLOCKING': 0,
                                'ABSENT': 0,
                                'WORK_HOURS': 0,
                                'OT_HOURS': 0,
                                'LVE DAYS': 0,
                                'MEAL': 0,
                                'MANUAL OT': 0
                            }

                    summary_table = [
                        ["LATE IN", "", monthly_counters['LATE_IN'], "ABSENT", monthly_counters['ABSENT'], "",
                         "OT 1.5HRS",
                         monthly_counters['OT1.5']],
                        ["EARLY OUT", "", monthly_counters['EARLY_OUT'], "LVE DAYS", monthly_counters['LVE DAYS'], "",
                         "OT 2.0HRS", monthly_counters['OT2.0']],
                        ["FORGOT CLOCKING", "", monthly_counters['FORGOT_CLOCKING'], "MEAL", monthly_counters["MEAL"],
                         "",
                         "OT 3.0HRS", monthly_counters['OT3.0'], "", "", "", "", "", "", "", "",
                         monthly_counters['OT_HOURS']],
                        ["MANUAL CLOCKING", "", "", "", "", "", "MANUAL OT HRS",
                         round(monthly_counters['MANUAL_OT'] / 60, 2)],
                        ["", "", "", "", "", "", "CANNOT OT", ""]
                    ]

                    employee_info = [
                        ["NAME", group.iloc[0]['Name'], "", "", "", "", "ID", emp_id, "", "", "SHIFT",
                         group.iloc[0]['SHIFT']]
                    ]

                    logging.info(f"💾 Generating {emp_id}-{group.iloc[0]['Name']} Employee Report...")

                    final_df = pd.DataFrame(final_rows)
                    source_sheet = wb_dept['Sheet1']
                    new_sheet = wb_dept.copy_worksheet(source_sheet)
                    new_sheet.title = str(emp_id)[:31]  # 限制 Sheet 名稱最多 31 字元

                    source_sheet_master = wb_master.active
                    # pd.DataFrame().to_excel(writer, sheet_name=sheet_name, index=False)
                    # ws = writer.book[sheet_name]

                    # 寫入 summary_table
                    for r_idx, row in enumerate(summary_table, 1):
                        for c_idx, value in enumerate(row, 2):
                            new_sheet.cell(row=r_idx, column=c_idx, value=value)

                    # 寫入 employee_info
                    for r_idx, row in enumerate(employee_info, len(summary_table) + 4):
                        for c_idx, value in enumerate(row, 1):
                            new_sheet.cell(row=r_idx, column=c_idx, value=value)

                    # 寫入 final_df
                    for r_idx, row in enumerate(dataframe_to_rows(final_df, index=False, header=False),
                                                start=len(summary_table) + len(employee_info) + 6):
                        for c_idx, value in enumerate(row, 1):
                            new_sheet.cell(row=r_idx, column=c_idx, value=value)

                    # 動態尋找含 "Summary up to" 的 summary 行，加黃底 & 粗體
                    for row in new_sheet.iter_rows(min_row=1, max_row=new_sheet.max_row):
                        for cell in row:
                            if cell.value and isinstance(cell.value, str) and "Summary up to" in cell.value:
                                for summary_cell in row:
                                    summary_cell.fill = yellow_fill
                                    summary_cell.font = Font(bold=True, size=16)
                                break

                    # 動態尋找 'WORK' 欄的位置，對Public Holiday行，加粉色底
                    header = [cell.value for cell in new_sheet[11]]
                    if 'DAY TYPE' in header:
                        day_type_col_index = header.index('DAY TYPE') + 1

                        # 從第2列開始逐列檢查
                        for row in new_sheet.iter_rows(min_row=1, max_row=new_sheet.max_row):
                            work_cell = row[day_type_col_index - 1]
                            if work_cell.value == 'PH':
                                for cell in row:
                                    cell.fill = pink_fill

                    logging.info(f"💾 Generating {emp_id}-{group.iloc[0]['Name']} Master Report...")
                    # 動態找出欄位位置
                    header = [cell.value for cell in source_sheet_master[1]]

                    # 取得對應欄位索引
                    stat_columns = {
                        'OT1.5': get_col_index('OT 1.5'),
                        'OT2.0': get_col_index('OT 2.0'),
                        'OT3.0': get_col_index('OT 3.0'),
                        'MANUAL_OT': get_col_index('MANUAL OT'),
                        'ABSENT': get_col_index('ABS'),
                        'MEAL': get_col_index('MEAL'),
                        'LVE DAYS': get_col_index('LVE DAYS'),
                        'CANNOT_OT': get_col_index('CANNOT OT'),
                        'LATE_IN': get_col_index('LATE IN'),
                        'EARLY_OUT': get_col_index('EARLY OUT'),
                        'FINAL_OT1.5': get_col_index('FINAL OT 1.5')
                    }

                    # 新增欄位：員工基本資料對應的欄位位置（也從 header 裡找）
                    basic_info_columns = {
                        'Type': get_col_index('Type'),
                        'Employee ID': get_col_index('Employee ID'),
                        'Name': get_col_index('Name (EN)'),
                        'Department': get_col_index('Department'),
                        'Shift': get_col_index('Shift'),
                        'On board date': get_col_index('On board date'),
                        'Leave date': get_col_index('Leave date[YYMMDD]')
                    }

                    monthly_counters['MANUAL_OT'] = round(monthly_counters['MANUAL_OT'] / 60, 2)
                    monthly_counters['FINAL_OT1.5'] = monthly_counters['OT1.5'] + monthly_counters['MANUAL_OT']

                    emp_info_row = employee_df[employee_df['Employee ID'] == emp_id]

                    if emp_info_row.empty:
                        logging.warning(
                            f"⚠️ Employee ID: {emp_id} can't be found in masterdata.xlsx --> Employee sheet")

                    else:
                        # 把該筆 Series 資料轉成字典
                        emp_info = emp_info_row.iloc[0].to_dict()

                        # 用來動態新增新列
                        target_row = find_first_empty_row(source_sheet_master, basic_info_columns['Employee ID'])

                        # 寫入員工基本資料
                        for field, col_index in basic_info_columns.items():
                            if col_index is not None:
                                source_sheet_master.cell(row=target_row, column=col_index + 1,
                                                         value=emp_info.get(field, ''))

                        # 寫入統計資料
                        for field, col_index in stat_columns.items():
                            if col_index is not None:
                                source_sheet_master.cell(row=target_row, column=col_index + 1,
                                                         value=monthly_counters.get(field, 0))

                logging.info("💾 Saving in progress...Please be patient")
                # 刪除原本的模板頁
                if "Sheet1" in wb_dept.sheetnames:
                    if len(wb_dept.sheetnames) > 1:
                        std = wb_dept["Sheet1"]
                        wb_dept.remove(std)

                wb_dept.save(DEPARTMENT_OUTPUT_PATH)
                wb_master.save(MASTER_OUTPUT_PATH)

        except Exception as e:
            logging.error("❌ An error occurred while executing the program！")
            logging.error(traceback.format_exc())
            messagebox.showerror("Error", f"An error occurred：{str(e)}\nPlease check the log file！")

# 完成後，彈窗提示
root = tk.Tk()
root.withdraw()
messagebox.showinfo("Completion Notification", "Reports have been successfully generated in the output folder！")
