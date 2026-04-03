"""Calculates total volume of all walls in the model."""

from Autodesk.Revit import DB

doc = __revit__.ActiveUIDocument.Document

wall_collector = DB.FilteredElementCollector(doc)\
    .OfCategory(DB.BuiltInCategory.OST_Walls)\
        .WhereElementIsNotElementType()
        
total_volume = 0.0

for wall in wall_collector:
    vol_param = wall.get_Parameter(DB.BuiltInParameter.HOST_VOLUME_COMPUTED)
    if vol_param:
        total_volume += vol_param.AsDouble()
        
print("Total Volume of Walls: {}".format(total_volume))