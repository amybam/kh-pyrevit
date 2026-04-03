"""Draws Revision Cloud around selected elements"""
# pylint: disable=import-error
import math
from pyrevit import revit, DB, script, forms

doc = revit.doc
uidoc = revit.uidoc
active_view = doc.ActiveView
output = script.get_output()

PADDING = 0.25  # feet of padding around elements


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def get_selected_elements():
    """Return list of selected elements, or None if nothing selected."""
    sel_ids = uidoc.Selection.GetElementIds()
    if not sel_ids:
        return None
    return [doc.GetElement(eid) for eid in sel_ids]


def get_latest_revision():
    """Return the ElementId of the latest revision, or None."""
    revisions = (
        DB.FilteredElementCollector(doc)
        .OfClass(DB.Revision)
        .WhereElementIsNotElementType()
        .ToElements()
    )
    if not revisions:
        return None
    return sorted(revisions, key=lambda r: r.SequenceNumber)[-1].Id


def _points_from_geom(geom_elem):
    """Recursively extract 2D points from geometry elements."""
    pts = []
    for geom_obj in geom_elem:
        if isinstance(geom_obj, DB.GeometryInstance):
            pts.extend(_points_from_geom(geom_obj.GetInstanceGeometry()))
        elif isinstance(geom_obj, DB.Solid):
            for edge in geom_obj.Edges:
                curve = edge.AsCurve()
                pts.append((curve.GetEndPoint(0).X, curve.GetEndPoint(0).Y))
                pts.append((curve.GetEndPoint(1).X, curve.GetEndPoint(1).Y))
        elif isinstance(geom_obj, DB.Curve):
            pts.append((geom_obj.GetEndPoint(0).X, geom_obj.GetEndPoint(0).Y))
            pts.append((geom_obj.GetEndPoint(1).X, geom_obj.GetEndPoint(1).Y))
        elif isinstance(geom_obj, DB.Line):
            pts.append((geom_obj.GetEndPoint(0).X, geom_obj.GetEndPoint(0).Y))
            pts.append((geom_obj.GetEndPoint(1).X, geom_obj.GetEndPoint(1).Y))
        elif isinstance(geom_obj, DB.PolyLine):
            for coord in geom_obj.GetCoordinates():
                pts.append((coord.X, coord.Y))
    return pts


def collect_geometry_points(elements, view):
    """Extract 2D points from actual visible geometry of elements.

    Uses view-specific geometry for tight bounds. Falls back to
    bounding box corners if geometry extraction yields nothing.
    """
    opts = DB.Options()
    opts.View = view
    opts.ComputeReferences = False

    all_points = []
    for elem in elements:
        elem_pts = []
        try:
            geom = elem.get_Geometry(opts)
            if geom:
                elem_pts = _points_from_geom(geom)
        except Exception:
            pass

        # Fall back to bounding box if geometry extraction failed
        if not elem_pts:
            bb = elem.get_BoundingBox(view)
            if bb:
                min_pt, max_pt = bb.Min, bb.Max
                elem_pts = [
                    (min_pt.X, min_pt.Y), (max_pt.X, min_pt.Y),
                    (max_pt.X, max_pt.Y), (min_pt.X, max_pt.Y),
                ]
        all_points.extend(elem_pts)
    return all_points


def cross(o, a, b):
    """2D cross product of vectors OA and OB."""
    return (a[0] - o[0]) * (b[1] - o[1]) - (a[1] - o[1]) * (b[0] - o[0])


def convex_hull(points):
    """Compute convex hull using Andrew's monotone chain algorithm.

    Returns vertices in counter-clockwise order.
    """
    pts = sorted(set(points))
    if len(pts) <= 1:
        return pts

    # Build lower hull
    lower = []
    for p in pts:
        while len(lower) >= 2 and cross(lower[-2], lower[-1], p) <= 0:
            lower.pop()
        lower.append(p)

    # Build upper hull
    upper = []
    for p in reversed(pts):
        while len(upper) >= 2 and cross(upper[-2], upper[-1], p) <= 0:
            upper.pop()
        upper.append(p)

    # Remove last point of each half because it's repeated
    return lower[:-1] + upper[:-1]


def offset_polygon(hull, distance):
    """Offset a convex polygon outward by the given distance."""
    n = len(hull)
    if n < 3:
        return hull

    offset_pts = []
    for i in range(n):
        # Previous, current, next points
        prev = hull[(i - 1) % n]
        curr = hull[i]
        nxt = hull[(i + 1) % n]

        # Edge vectors
        e1 = (curr[0] - prev[0], curr[1] - prev[1])
        e2 = (nxt[0] - curr[0], nxt[1] - curr[1])

        # Outward normals (rotate 90 degrees clockwise for CCW polygon = left normal)
        def normalize(v):
            length = math.sqrt(v[0] ** 2 + v[1] ** 2)
            if length < 1e-10:
                return (0.0, 0.0)
            return (v[0] / length, v[1] / length)

        n1 = normalize((-e1[1], e1[0]))
        n2 = normalize((-e2[1], e2[0]))

        # Average normal at vertex
        avg = normalize((n1[0] + n2[0], n1[1] + n2[1]))

        # Offset vertex
        offset_pts.append((curr[0] + avg[0] * distance,
                           curr[1] + avg[1] * distance))

    return offset_pts



# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

elements = get_selected_elements()
if not elements:
    forms.alert("No elements selected. Please select elements first.",
                exitscript=True)

revision_id = get_latest_revision()
if not revision_id:
    forms.alert("No revisions found in the project. "
                "Please add a revision first.",
                exitscript=True)

corners = collect_geometry_points(elements, active_view)
if not corners:
    forms.alert("Could not get geometry for the selected elements.",
                exitscript=True)

hull = convex_hull(corners)

# Fall back to bounding rectangle if hull is degenerate
if len(hull) < 3:
    xs = [p[0] for p in corners]
    ys = [p[1] for p in corners]
    min_x, max_x = min(xs), max(xs)
    min_y, max_y = min(ys), max(ys)
    hull = [(min_x, min_y), (max_x, min_y),
            (max_x, max_y), (min_x, max_y)]

hull.reverse()  # CW winding so revision cloud bumps face outward
hull = offset_polygon(hull, PADDING)

from System.Collections.Generic import List
curves = List[DB.Curve]()
for i in range(len(hull)):
    x1, y1 = hull[i]
    x2, y2 = hull[(i + 1) % len(hull)]
    p1 = DB.XYZ(x1, y1, 0.0)
    p2 = DB.XYZ(x2, y2, 0.0)
    curves.Add(DB.Line.CreateBound(p1, p2))

with revit.Transaction("Create Revision Cloud"):
    cloud = DB.RevisionCloud.Create(doc, active_view,
                                    revision_id, curves)

print("Revision cloud created around {} elements.".format(len(elements)))
