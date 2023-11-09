# Import necessary Revit API modules
import clr
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, BuiltInParameter

# Access the active document in Revit
doc = __revit__.ActiveUIDocument.Document

# Get all rooms in the project
rooms = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rooms).WhereElementIsNotElementType()

# Loop through each room and print area, name, and number
for room in rooms:
    area = room.get_Parameter(BuiltInParameter.ROOM_AREA).AsDouble()
    name = room.get_Parameter(BuiltInParameter.ROOM_NAME).AsString()
    number = room.get_Parameter(BuiltInParameter.ROOM_NUMBER).AsString()
    
    print("Room Name: {}, Room Number: {}, Area: {:.2f} sq.ft.".format(name, number, area))
