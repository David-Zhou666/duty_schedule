import openpyxl

file_path = 'src/main/resources/duty_temple.xlsx'
wb = openpyxl.load_workbook(file_path)

for sheet_name in ['一班']:  # 只检查第一个sheet
    ws = wb[sheet_name]
    print(f"Sheet: {sheet_name}\n")
    print("完整数据（前13行）:")
    print(f"{'行':<3} | {'[0]':<8} | {'[1]':<10} | {'[2]':<6} | {'[3]':<4} | {'[4]':<4} | {'[5]':<4} | {'[6]':<4} | {'[7]':<4}")
    print("-" * 80)
    
    for row_idx, row in enumerate(ws.iter_rows(max_row=13, values_only=True), 1):
        col0 = str(row[0]) if row[0] else ""
        col1 = str(row[1]) if row[1] else ""
        col2 = str(row[2]) if row[2] else ""
        col3 = str(row[3]) if row[3] else ""
        col4 = str(row[4]) if row[4] else ""
        col5 = str(row[5]) if row[5] else ""
        col6 = str(row[6]) if row[6] else ""
        col7 = str(row[7]) if row[7] else ""
        print(f"{row_idx:<3} | {col0:<8} | {col1:<10} | {col2:<6} | {col3:<4} | {col4:<4} | {col5:<4} | {col6:<4} | {col7:<4}")
