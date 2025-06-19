import pandas as pd
from datetime import datetime, timedelta

# 設定固定參數
WORK_START = datetime.strptime('08:00:00', '%H:%M:%S').time()
WORK_END = datetime.strptime('18:00:00', '%H:%M:%S').time()
WEEK_START = 0  # 週一是第0天
WEEK_END = 6    # 週日是第6天

# 匯入資料
master_df = pd.read_excel('./data/master.xlsx')
leave_df = pd.read_excel('./data/leave.xlsx')

# 基本清理
master_df['Date'] = pd.to_datetime(master_df['Date'])
master_df.sort_values(['Employee ID', 'Date'], inplace=True)

# 增加 LATE_MIN 和 EARLY_MIN
def calc_late(clock_in, base_time=WORK_START):
    if clock_in == "-":
        return "-"
    t = pd.to_datetime(clock_in).time()
    if t > base_time:
        return (datetime.combine(datetime.min, t) - datetime.combine(datetime.min, base_time)).seconds // 60
    else:
        return "-"

def calc_early(clock_out, base_time=WORK_END):
    if clock_out == "-":
        return "-"
    t = pd.to_datetime(clock_out).time()
    if t < base_time:
        return (datetime.combine(datetime.min, base_time) - datetime.combine(datetime.min, t)).seconds // 60
    else:
        return "-"

master_df['LATE_MIN'] = master_df['Clock-in'].apply(lambda x: calc_late(x))
master_df['EARLY_MIN'] = master_df['Clock-out'].apply(lambda x: calc_early(x))

# 周六周日標註
master_df['DAY'] = master_df['Date'].dt.day_name()
master_df['DAY_NUM'] = master_df['Date'].dt.weekday

# 處理忘記打卡
master_df['FORGOT_CLOCKING'] = master_df.apply(
    lambda x: 1 if (x['Clock-in'] == "-" or x['Clock-out'] == "-") and not(x['DAY_NUM'] == 5 or x['DAY_NUM'] == 6) else 0,
    axis=1
)

# 計算缺勤（ABSENT）
def is_absent(row):
    if row['FORGOT_CLOCKING'] == 1:
        return 1
    return 0
master_df['ABSENT'] = master_df.apply(is_absent, axis=1)

# 每天工時 WORK
def calc_work_hours(in_time, out_time):
    if in_time == "-" or out_time == "-":
        return 0
    t_in = pd.to_datetime(in_time)
    t_out = pd.to_datetime(out_time)
    work_minutes = (t_out - t_in).seconds / 3600
    return round(work_minutes, 2)

master_df['WORK'] = master_df.apply(lambda x: calc_work_hours(x['Clock-in'], x['Clock-out']), axis=1)

# 合併請假資料 (簡單版：直接貼在對應日)
def map_leave(date, emp_id):
    record = leave_df[(leave_df['Employee ID'] == emp_id) &
                      (leave_df['Start Date'] <= date) &
                      (leave_df['End Date'] >= date)]
    if not record.empty:
        return record.iloc[0]['Leave Type']
    else:
        return "-"

master_df['LEAVE'] = master_df.apply(lambda x: map_leave(x['Date'], x['Employee ID']), axis=1)

# 計算 LIVE DAYS（出勤天數）
# master_df['LIVE DAYS'] = master_df.apply(
#     lambda x: 1 if (x['DAY TYPE'] == 'WORK') and (x['FORGOT_CLOCKING'] == 0) else 0,
#     axis=1
# )


master_df['SHIFT'] = 'B1'  # 固定
master_df['MEAL'] = 3  # 固定
master_df['OT'] = 0.0  # 先設成0

def auto_day_type(row):
    if row['DAY_NUM'] == 6:
        return 'REST'
    elif row['DAY_NUM'] == 5:
        return 'OFF'  # 週六排休
    else:
        return 'WORK'

master_df['DAY TYPE'] = master_df.apply(auto_day_type, axis=1)

# 生成報表（含周總結）
final_rows = []
week_stats = {
    'LATE_IN': 0,
    'EARLY_OUT': 0,
    'FORGOT_CLOCKING': 0,
    'ABSENT': 0,
    'LIVE_DAYS': 0,
    'OT_HOURS': 0,
}

for idx, row in master_df.iterrows():
    final_rows.append(row)

    if row['DAY_NUM'] == WEEK_END:  # 遇到周日就插入統計
        week_summary = {
            'Employee ID': '',
            'Name': '',
            'Company / Department': '',
            'Sex': '',
            'Date': '',
            'Day': f"Summary up to {row['Date'].strftime('%Y-%m-%d')}",
            'Clock-in': '',
            'Clock-out': '',
            'LATE_MIN': week_stats.get('LATE_IN', '-'),
            'EARLY_MIN': week_stats.get('EARLY_OUT', '-'),
            'FORGOT_CLOCKING': week_stats.get('FORGOT_CLOCKING', 0),
            'ABSENT': week_stats.get('ABSENT', 0),
            'WORK': '',
            'LEAVE': '',
            'SHIFT': '',
            'MEAL': '',
            'OT': week_stats['OT_HOURS'],
            'DAY TYPE': ''
        }
        final_rows.append(pd.Series(week_summary))
        # 重置週統計
        week_stats = dict.fromkeys(week_stats, 0)

    # 累加統計
    week_stats['LATE_IN'] += row['LATE_MIN'] if isinstance(row['LATE_MIN'], int) else 0
    week_stats['EARLY_OUT'] += row['EARLY_MIN'] if isinstance(row['EARLY_MIN'], int) else 0
    week_stats['FORGOT_CLOCKING'] += row['FORGOT_CLOCKING']
    week_stats['ABSENT'] += row['ABSENT']
#   week_stats['LIVE_DAYS'] += row['LIVE_DAYS']
    week_stats['OT_HOURS'] += row['OT']

# 出口成EXCEL
final_df = pd.DataFrame(final_rows)
final_df.to_excel('./output/考勤報表_完成版.xlsx', index=False)

print("✅ 報表生成完成！")

