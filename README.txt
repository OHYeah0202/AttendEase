考勤統計自動化腳本使用說明
版本：2.2
作者：YTHsieh
日期：2025-05-20

【📌 目的】  
本腳本自動化統計公司員工的每月考勤資料，根據部門自動區分直接人員 (DL) 與間接人員 (IDL)，並統計加班時數（OT 1.5、OT 2.0、OT 3.0）、缺勤、遲到早退等項目，同時將資料寫入標準模板報表中。

【使用方式】
1. 請將 Excel 檔案命名 masterdata.xlsx 並命名 Excel 表為：
   ➤ `Employee`: 所有員工資料
   ➤ `Holiday`: 公共假日及周六不加班資料 
   ➤ `Attendance`: 考勤打卡資料 
   ➤ `Leave`: 請假記錄資料
   ➤ `Meal`: 餐費津貼資料
   並放置於 `data/`資料夾内。

2. 確認範本檔 `employee_report_template.xlsx` `master_report_template.xlsx` 存在於 `template/` 資料夾内，樣式將會自動複製到新報表中。

3. 執行主腳本檔案：
   ➤ `main_v2.2.exe`

4. 程式將依據以下規則分類、計算與輸出考勤結果，並將統計結果寫入報表中。

【規則說明】
1. 🕒 **出勤時段區分**
   - A1：上班時間為 07:50～18:40（含 1 小時 50 分鐘休息）
   -  A2: 上班時間為 07:50～18:40（含 1 小時 50 分鐘休息），周五上班時間為 07:50 ~ 19:40（含 1 小時 50 分鐘休息）
   - B1: 間接人員：上班時間為 08:00～18:00（含 1 小時休息）
   - 透過 masterdata.xlsx --> Employee Excel 表 `Shift` 欄位判斷

2. **加班時數計算（以每 30 分鐘為單位，不足者不計）：**
   - OT 1.5：平日的加班時間
   - OT 2.0：星期日、或國定假日前 8 小時
   - OT 3.0：國定假日超過 8 小時部分

3. 🚫 **缺勤與異常類型統計**  
   - 曠職（ABSENT）
   - 請假（LVE DAYS）
   - 遲到（LATE IN） / 早退（EARLY OUT）
   - 忘打卡、不可加班等（FORGOT CLOCKING、CANNOT OT）

4. 🧾 **報表輸出格式**  
   程式將於 `output/` 資料夾內輸出以下兩份 Excel 報表：
   ➤ `[Month]_Employees_Report.xlsx`：個別員工考勤統計  
   ➤ `Master_Report.xlsx`：總表，含員工基本資料與統計資料 
   - 所有資料將自動套用範本樣式  

【⚠️ 注意事項】 
- 請勿修改 Excel 模板格式與欄位名稱，以避免分析錯誤。
- 請假與加班記錄需正確填寫於對應檔案中，避免漏算。
- 加班時數僅供人資部門內部參考，實際薪資請依公司核算標準為準。
- 若需加入其他考勤類別或修改計算邏輯，請聯絡開發者進行擴充。

【📨 聯絡方式】
如有任何問題，歡迎聯絡 YTHsieh
Email: ythsieh@altekmed.com.my

Attendance Statistics Automation Script User Guide

Version: 2.2
Author: YTHsieh
Date: 2025-05-20


📌 Purpose
This script automates the monthly attendance statistics for company employees. It classifies employees into Direct Labor (DL) and Indirect Labor (IDL) based on department, calculates overtime hours (OT 1.5, OT 2.0, OT 3.0), absences, lateness/early leave, and writes the results into standardized report templates.


How to Use

1. Prepare a single Excel file named masterdata.xlsx, and make sure it contains the following sheets:
➤ Employee: Employee master data
➤ Holiday: Public holidays and Saturdays without overtime
➤ Attendance: Attendance clock-in/out records
➤ Leave: Leave records
➤ Meal: Meal allowance data
Place this file inside the data/ folder.

2. Ensure the following template files exist in the template/ folder:
➤ employee_report_template.xlsx
➤ master_report_template.xlsx
These templates will be used to style the output reports.

3. Run the main script executable:
➤ main_v2.2.exe

4. The program will classify, calculate, and output attendance results based on the following rules, and write them into formatted reports.


Rules Description

1. 🕒 Shift Classification
· A1: 07:50–18:40 (including 1 hour 50 minutes break)
· A2: 07:50–18:40 (including 1 hour 50 minutes break); on Fridays: 07:50–19:40
· B1: 08:00–18:00 (including 1 hour break)
Shift type is determined from the Shift column in the Employee sheet.

2. ⏱️ Overtime Calculation (in 30-minute blocks; under 30 minutes not counted):
· OT 1.5: Weekday overtime
· OT 2.0: Sundays and the first 8 hours on public holidays
· OT 3.0: Hours exceeding 8 on public holidays

3. 🚫 Absence and Exception Tracking
· ABSENT: Unexcused absence
· LVE DAYS: Official leave
· LATE IN / EARLY OUT: Lateness and early departure
· FORGOT CLOCKING / CANNOT OT: Missed punch-in/out, overtime not allowed

4. 🧾 Report Output
The following Excel reports will be generated in the output/ folder:
➤ [Month]_Employees_Report.xlsx: Individual employee attendance summary
➤ Master_Report.xlsx: Consolidated report including employee data and statistics
All data will be automatically styled using the provided templates


⚠️ Notes

· Do not alter the Excel template column names to avoid processing errors.

· Ensure leave and overtime records are filled correctly in the corresponding sheets to prevent omissions.

· Overtime statistics are for internal HR reference only; actual payroll should follow company policies.

· For adding new attendance categories or modifying calculation logic, please contact the developer.


📨 Contact
For any inquiries, feel free to contact YTHsieh
Email: ythsieh@altekmed.com.my