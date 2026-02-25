# -*- coding: utf-8 -*-
"""验证生成的排班表"""
import openpyxl

wb = openpyxl.load_workbook('duty_schedule_result.xlsx')
ws = wb.active

print('排班表内容:')
print('='*90)
for row in ws.iter_rows(values_only=True):
    print(f'{row[0]:<8} {row[1]:<12} {row[2]:<12} {row[3]:<12} {row[4]:<12} {row[5]:<12} {row[6]:<12}')
print('='*90)
print()
print('排版说明:')
print('1. 每列显示具体学生姓名（如：李伟、刘峰等）')
print('2. 不使用"男生"、"女生"字样，而是直接显示姓名')
print('3. 每天需要4名男生（男1-男4）和2名女生（女1-女2）')
print('4. 每名同学一周不超过5次值日')
