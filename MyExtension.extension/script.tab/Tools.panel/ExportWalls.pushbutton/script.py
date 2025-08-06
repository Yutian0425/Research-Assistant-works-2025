# -*- coding: utf-8 -*-

from pyrevit import revit, DB
import os
import sys
import json
import clr

clr.AddReference('RevitAPI')

from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory

# 使用 pyRevit 提供的 doc（不使用 DocumentManager）
doc = revit.doc

# 获取当前脚本所在的目录
current_dir = os.path.dirname(__file__)

# 外部 Python 脚本路径
external_python_script = os.path.join(current_dir, 'export_walls.py')

# 指定 Python 解释器路径（如果 pyRevit 环境中没有 sys.executable）
python_path = r'C:\Users\admin\AppData\Local\Programs\Python\Python313\python.exe'
python_executable = sys.executable if sys.executable else python_path

# 获取墙体数据
wall_data = []
walls = FilteredElementCollector(doc) \
    .OfCategory(BuiltInCategory.OST_Walls) \
    .WhereElementIsNotElementType() \
    .ToElements()

for wall in walls:
    wall_data.append({
        'name': wall.Name,
        'id': str(wall.Id)
    })

import tempfile

# 写入 JSON
json_data = json.dumps(wall_data)
temp_file = os.path.join(tempfile.gettempdir(), 'walls_data.json')
with open(temp_file, 'w') as f:
    f.write(json_data)

# 调用外部 Python 脚本
os.system('"{}" "{}" "{}"'.format(python_executable, external_python_script, temp_file))
