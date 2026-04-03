"""Draws Revision Cloud around selected elements"""
# pylint: disable=import-error
import math
import os
import clr

clr.AddReference('System.Drawing')
from System.Drawing import Bitmap  # noqa: E402
from System.Collections.Generic import List  # noqa: E402

from pyrevit import revit, DB, script, forms  # noqa: E402

doc = revit.doc
uidoc = revit.uidoc
active_view = doc.ActiveView
output = script.get_output()

PADDING = 0.25  # feet of padding around elements
IMAGE_SIZE = 800  # pixels for the exported image


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


# ---------------------------------------------------------------------------
# Vision-based boundary detection
# ---------------------------------------------------------------------------

def _get_bbox_region(elements, view, margin=2.0):
    """Get the combined bounding box region of elements with margin."""
    bbs = [e.get_BoundingBox(view) for e in elements]
    bbs = [bb for bb in bbs if bb is not None]
    if not bbs:
        return None
    return (
        min(bb.Min.X for bb in bbs) - margin,
        min(bb.Min.Y for bb in bbs) - margin,
        max(bb.Max.X for bb in bbs) + margin,
        max(bb.Max.Y for bb in bbs) + margin,
    )


def _export_isolated_image(elements, view, region):
    """Isolate elements, set crop to known region, export image.

    Returns the path to the exported PNG, or None on failure.
    """
    rgn_min_x, rgn_min_y, rgn_max_x, rgn_max_y = region
    elem_ids = List[DB.ElementId]([e.Id for e in elements])

    # Save view state
    was_crop_active = view.CropBoxActive
    was_crop_visible = view.CropBoxVisible
    orig_crop = view.CropBox

    # Set crop to our known coordinate region
    t = DB.Transaction(doc, "Temp crop for vision export")
    t.Start()
    try:
        new_crop = view.CropBox
        new_crop.Min = DB.XYZ(rgn_min_x, rgn_min_y, new_crop.Min.Z)
        new_crop.Max = DB.XYZ(rgn_max_x, rgn_max_y, new_crop.Max.Z)
        view.CropBox = new_crop
        view.CropBoxActive = True
        view.CropBoxVisible = False
        t.Commit()
    except Exception:
        if t.HasStarted() and not t.HasEnded():
            t.RollBack()
        return None

    # Isolate only selected elements (no transaction needed)
    view.IsolateElementsTemporary(elem_ids)
    doc.Regenerate()

    # Export image
    temp_dir = os.environ.get('TEMP', os.environ.get('TMP', 'C:\\Temp'))
    temp_base = "revcloud_vision_temp"

    exported = None
    try:
        img_opts = DB.ImageExportOptions()
        img_opts.ExportRange = DB.ExportRange.CurrentView
        img_opts.ZoomType = DB.ZoomFitType.FitToPage
        img_opts.PixelSize = IMAGE_SIZE
        img_opts.FilePath = os.path.join(temp_dir, temp_base)
        img_opts.HLRandWFViewsFileType = DB.ImageFileType.PNG
        img_opts.ShadowViewsFileType = DB.ImageFileType.PNG
        img_opts.ImageResolution = DB.ImageResolution.DPI_150
        doc.ExportImage(img_opts)
    except Exception:
        pass

    # Restore view state immediately
    view.DisableTemporaryViewMode(
        DB.TemporaryViewMode.TemporaryHideIsolate)

    t2 = DB.Transaction(doc, "Restore crop after vision export")
    t2.Start()
    try:
        view.CropBox = orig_crop
        view.CropBoxActive = was_crop_active
        view.CropBoxVisible = was_crop_visible
        t2.Commit()
    except Exception:
        if t2.HasStarted() and not t2.HasEnded():
            t2.RollBack()

    # Locate the exported file (Revit may append view name)
    for f in os.listdir(temp_dir):
        if f.startswith(temp_base) and f.lower().endswith('.png'):
            exported = os.path.join(temp_dir, f)
            break

    return exported


def _scan_boundary_pixels(image_path):
    """Scan exported image and return boundary pixel coordinates.

    Uses row/column edge sweeps for efficiency: for each row find the
    leftmost and rightmost non-background pixel, and for each column
    find the topmost and bottommost.  These boundary pixels trace the
    visible outline of the elements.
    """
    bmp = Bitmap(image_path)
    w, h = bmp.Width, bmp.Height
    bg = bmp.GetPixel(0, 0)  # assume top-left corner is background

    def is_bg(px, py):
        c = bmp.GetPixel(px, py)
        return c.R == bg.R and c.G == bg.G and c.B == bg.B

    boundary = []  # list of (pixel_x, pixel_y)
    row_step = max(1, h // 150)
    col_step = max(1, w // 150)

    # Sweep each row: find leftmost and rightmost visible pixels
    for py in range(0, h, row_step):
        # Left edge
        for px in range(0, w):
            if not is_bg(px, py):
                boundary.append((px, py))
                break
        # Right edge
        for px in range(w - 1, -1, -1):
            if not is_bg(px, py):
                boundary.append((px, py))
                break

    # Sweep each column: find topmost and bottommost visible pixels
    for px in range(0, w, col_step):
        # Top edge
        for py in range(0, h):
            if not is_bg(px, py):
                boundary.append((px, py))
                break
        # Bottom edge
        for py in range(h - 1, -1, -1):
            if not is_bg(px, py):
                boundary.append((px, py))
                break

    bmp.Dispose()
    return boundary, w, h


def collect_points_via_vision(elements, view):
    """Detect tight element boundaries by exporting a view image,
    scanning for visible pixels, and mapping back to model coordinates.

    Falls back to bounding box corners on failure.
    """
    region = _get_bbox_region(elements, view)
    if not region:
        return _collect_bbox_corners(elements, view)

    rgn_min_x, rgn_min_y, rgn_max_x, rgn_max_y = region

    image_path = _export_isolated_image(elements, view, region)
    if not image_path or not os.path.exists(image_path):
        print("Vision export failed, falling back to bounding boxes.")
        return _collect_bbox_corners(elements, view)

    try:
        boundary_pixels, img_w, img_h = _scan_boundary_pixels(image_path)
    except Exception:
        print("Image scan failed, falling back to bounding boxes.")
        return _collect_bbox_corners(elements, view)
    finally:
        try:
            os.remove(image_path)
        except Exception:
            pass

    if not boundary_pixels:
        print("No visible pixels found, falling back to bounding boxes.")
        return _collect_bbox_corners(elements, view)

    # Map pixel coordinates to model coordinates
    # Image (0,0) = top-left = (rgn_min_x, rgn_max_y)
    # Image (w,h) = bottom-right = (rgn_max_x, rgn_min_y)
    rgn_w = rgn_max_x - rgn_min_x
    rgn_h = rgn_max_y - rgn_min_y

    model_points = []
    for px, py in boundary_pixels:
        mx = rgn_min_x + (px / float(img_w)) * rgn_w
        my = rgn_max_y - (py / float(img_h)) * rgn_h
        model_points.append((mx, my))

    print("Vision detected {} boundary points.".format(len(model_points)))
    return model_points


def _collect_bbox_corners(elements, view):
    """Fallback: collect 2D corners from bounding boxes."""
    points = []
    for elem in elements:
        bb = elem.get_BoundingBox(view)
        if bb is None:
            continue
        min_pt, max_pt = bb.Min, bb.Max
        points.append((min_pt.X, min_pt.Y))
        points.append((max_pt.X, min_pt.Y))
        points.append((max_pt.X, max_pt.Y))
        points.append((min_pt.X, max_pt.Y))
    return points


# ---------------------------------------------------------------------------
# Convex hull and offset
# ---------------------------------------------------------------------------

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

    def normalize(v):
        length = math.sqrt(v[0] ** 2 + v[1] ** 2)
        if length < 1e-10:
            return (0.0, 0.0)
        return (v[0] / length, v[1] / length)

    offset_pts = []
    for i in range(n):
        prev = hull[(i - 1) % n]
        curr = hull[i]
        nxt = hull[(i + 1) % n]

        e1 = (curr[0] - prev[0], curr[1] - prev[1])
        e2 = (nxt[0] - curr[0], nxt[1] - curr[1])

        # Outward normals for CW winding
        n1 = normalize((-e1[1], e1[0]))
        n2 = normalize((-e2[1], e2[0]))

        avg = normalize((n1[0] + n2[0], n1[1] + n2[1]))

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

# Use vision-based detection for tight boundaries
corners = collect_points_via_vision(elements, active_view)
if not corners:
    forms.alert("Could not detect element boundaries.",
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
