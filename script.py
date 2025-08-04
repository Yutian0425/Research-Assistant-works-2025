# pyRevit script: export wall names and IDs to Excel using xlwings

import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitServices')

from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, Element
from RevitServices.Persistence import DocumentManager

import xlwings as xw
import os

# Access active Revit document
doc = DocumentManager.Instance.CurrentDBDocument

# Collect all wall instances in the model
walls = FilteredElementCollector(doc) \
    .OfCategory(BuiltInCategory.OST_Walls) \
    .WhereElementIsNotElementType() \
    .ToElements()

# Create a new Excel workbook
wb = xw.Book()  # Opens a new Excel instance
ws = wb.sheets[0]
ws.name = "Revit_Walls"

# Write headers
ws.range("A1").value = ["Wall Name", "Element ID"]

# Write wall data
row = 2
for wall in walls:
    ws.range("A{}".format(row)).value = wall.Name
    ws.range("B{}".format(row)).value = str(wall.Id)
    row += 1

# Save file
save_path = os.path.expanduser("~\\Documents\\RevitWallsExport.xlsx")
wb.save(save_path)
wb.close()

# Notify user with corrected print statement
print("✅ Export completed. File saved to:\n{}".format(save_path))
