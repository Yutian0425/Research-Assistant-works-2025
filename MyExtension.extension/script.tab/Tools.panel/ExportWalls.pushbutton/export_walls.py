# -*- coding: utf-8 -*-

# External Python script: Read Revit data from JSON and export it to Excel using xlwings

import sys
import json
import xlwings as xw
import os

# 获取 pyRevit 脚本传递的 JSON 文件路径
json_file = sys.argv[1]

# 读取 JSON 数据
with open(json_file, 'r', encoding='utf-8') as f:
    wall_data = json.load(f)

# Create a new Excel workbook
wb = xw.Book()  # Opens a new Excel instance
ws = wb.sheets[0]
ws.name = "Revit_Walls"

# Write headers
ws.range("A1").value = ["Wall Name", "Element ID"]

# Write wall data
row = 2
for wall in wall_data:
    ws.range(f"A{row}").value = wall['name']
    ws.range(f"B{row}").value = wall['id']
    row += 1

# Save file to Desktop
save_path = r"C:\Users\admin\Desktop\RevitWallsExport.xlsx"  # 保存到桌面
wb.save(save_path)
wb.close()

# Ensure Excel closes
wb.app.quit()

print(f"✅ Export completed. File saved to: {save_path}")
