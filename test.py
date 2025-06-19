import pandas as pd
from datetime import datetime, timedelta
import numpy as np
import os
import math

# 基準設定
WORK_START = datetime.strptime('08:00:00', '%H:%M:%S').time()
WORK_END = datetime.strptime('18:00:00', '%H:%M:%S').time()
DL_WORK_START = datetime.strptime('07:50:00', '%H:%M:%S').time()
DL_WORK_END = datetime.strptime('18:40:00', '%H:%M:%S').time()
LUNCH_BREAK = 1  # 1 小時午餐
WEEK_START = 0  # 周一是第0天
WEEK_END = 6  # 周日是第6天
OT_15 = OT_20 = OT_30 = 0


# 讀取資料
master_df = pd.read_excel('./data/master.xlsx')
leave_df = pd.read_excel('./data/leave.xlsx')

# 格式轉換
master_df['Date'] = pd.to_datetime(master_df['Date'])
master_df.sort_values(['Employee ID', 'Date'], inplace=True)
leave_df['Start Date'] = pd.to_datetime(leave_df['Start Date'])
leave_df['End Date'] = pd.to_datetime(leave_df['End Date'])


# 計算欄位：LATE_MIN 和 EARLY_MIN
# def calc_late(clock_in, base_time=WORK_START):
#     if pd.isna(clock_in):
#         return "-"
#     t = pd.to_datetime(clock_in).time()
#     if t > base_time:
#         return (datetime.combine(datetime.min, t) - datetime.combine(datetime.min, base_time)).seconds // 60
#     else:
#         return "-"
#
# def calc_early(clock_out, base_time=WORK_END):
#     if pd.isna(clock_out):
#         return "-"
#     t = pd.to_datetime(clock_out).time()
#     if t < base_time:
#         return (datetime.combine(datetime.min, base_time) - datetime.combine(datetime.min, t)).seconds // 60
#     else:
#         return "-"
#
# master_df['LATE_MIN'] = master_df['Clock-in'].apply(lambda x: calc_late(x))
# master_df['EARLY_MIN'] = master_df['Clock-out'].apply(lambda x: calc_early(x))

# 忘記打卡
# master_df['FORGOT_CLOCKING'] = master_df.apply(
#     lambda row: 1 if pd.isna(row['Clock-in']) or pd.isna(row['Clock-out']) else 0, axis=1)

# 缺勤 (整天無打卡且非請假)
# master_df['ABS'] = master_df.apply(
#     lambda row: 1 if pd.isna(row['Clock-in']) and pd.isna(row['Clock-out']) else 0, axis=1)


# 出勤工時計算
def calc_work_hours(row):
    if row['Clock-in'] == "-" or row['Clock-out'] == "-":
        return 0

    start = WORK_START

    if datetime.strptime(row['Clock-in'], '%H:%M:%S').time() > WORK_START:
        start = datetime.strptime(row['Clock-in'], '%H:%M:%S').time()

    end = datetime.strptime(row['Clock-out'], '%H:%M:%S').time()
    hours = (datetime.combine(datetime.today(), end) - datetime.combine(datetime.today(), start)).seconds / 3600 - LUNCH_BREAK
    return max(hours, 0)

master_df['WORK_HOURS'] = master_df.apply(calc_work_hours, axis=1)

def calc_ot_mins(row):
    global OT_15, OT_20, OT_30

    if not "IDL" in row["Company / Department"]:
        work_hours = row["WORK_HOURS"]

        if row["DAY TYPE"] == "Holiday":
            ot_units = math.floor(work_hours * 60 / 30) * 0.5

            if ot_units > 8:
                OT_20 += 8
                OT_30 += ot_units - 8
            else:
                OT_20 += ot_units

            return ot_units

        elif row["Day"] == "Sun.":
            ot_units = math.floor(work_hours * 60 / 30) * 0.5
            OT_20 += ot_units
            return ot_units

        else:
            ot_hours = max(0, work_hours - 9) # 超出標準9小時才算加班
            ot_units = math.floor(ot_hours * 60 / 30) * 0.5 # 換成小數點
            OT_15 += ot_units
            return ot_units

master_df['OT'] = master_df.apply(calc_ot_mins, axis=1)
# 處理請假標記
master_df['LEAVE'] = None
for _, leave_row in leave_df.iterrows():
    mask = (master_df['Employee ID'] == leave_row['Employee ID']) & (master_df['Date'] >= leave_row['Start Date']) & (
                master_df['Date'] <= leave_row['End Date'])
    master_df.loc[mask, 'LEAVE'] = leave_row['Leave Type']

# 按員工分組產出
for emp_id, emp_data in master_df.groupby('Employee ID'):
    emp_name = emp_data['Name'].iloc[0]
    emp_data = emp_data.sort_values('Date')

    # 插入週小結
    emp_data['WeekNum'] = emp_data['Date'].dt.isocalendar().week
    weekly_summary = []
    for week, week_group in emp_data.groupby('WeekNum'):
        summary_row = {
            'DATE': f'Week {week} Summary',
            'LATE_MIN': week_group['LATE_MIN'].sum(),
            'EARLY_MIN': week_group['EARLY_MIN'].sum(),
            'FORGOT_CLOCKING': week_group['FORGOT_CLOCKING'].sum(),
            'ABS': week_group['ABS'].sum(),
            'WORK_HOURS': week_group['WORK_HOURS'].sum(),
            'LEAVE': f"AL:{(week_group['LEAVE'] == 'Annual Leave').sum()} SL:{(week_group['LEAVE'] == 'Sick Leave').sum()}"
        }
        weekly_summary.append(summary_row)

    # 合併小結
    summary_df = pd.DataFrame(weekly_summary)
    emp_final = pd.concat([emp_data, summary_df], ignore_index=True)

    # 輸出Excel
    os.makedirs('output', exist_ok=True)
    emp_final.to_excel(f'output/{emp_id}_{emp_name}_月報.xlsx', index=False)

print("✅ 進階版日資料 + 週小結 完成！")