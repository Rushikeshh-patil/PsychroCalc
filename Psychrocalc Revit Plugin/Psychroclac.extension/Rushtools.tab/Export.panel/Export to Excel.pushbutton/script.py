# Import necessary Revit API modules
import clr
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, BuiltInParameter
import csv

# Access the active document in Revit
doc = __revit__.ActiveUIDocument.Document

# Get all rooms in the project
rooms = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rooms).WhereElementIsNotElementType()

# Define the file path
file_path = "C:\\Users\\rush\\OneDrive - THERMA CORPORATION\\Desktop\\room_data.csv"

with open(file_path, mode='w') as file:
    writer = csv.writer(file)
    writer.writerow(["Room Name", "Room Number", "Area (sq.ft.)"])  # Write header row
    for room in rooms:
        area = room.get_Parameter(BuiltInParameter.ROOM_AREA).AsDouble()
        name = room.get_Parameter(BuiltInParameter.ROOM_NAME).AsString()
        number = room.get_Parameter(BuiltInParameter.ROOM_NUMBER).AsString()
        writer.writerow([name, number, round(area, 2)])

    print("Data exported to " + file_path)
