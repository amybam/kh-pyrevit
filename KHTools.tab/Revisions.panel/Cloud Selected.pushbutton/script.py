"""Draws Revision Cloud around selected elements"""
# pylint: disable=import-error
from collections import defaultdict
from pyrevit import revit, DB, script, forms
from System.Collections.Generic import List

doc = revit.doc
uidoc = revit.uidoc
active_view = doc.ActiveView
output = script.get_output()

PADDING_VIEW = 0.25  # feet (3 inches) for model views
PADDING_SHEET = 0.25 / 12.0  # feet (0.25 inches) for sheets

PADDING = PADDING_SHEET if isinstance(active_view, DB.ViewSheet) else PADDING_VIEW


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


def get_padded_bboxes(elements, view, padding):
    """Get axis-aligned bounding boxes for elements, expanded by padding.

    A tiny epsilon is added to prevent boxes from touching at only a corner,
    which would complicate boundary tracing.
    """
    eps = 1e-6
    boxes = []
    for elem in elements:
        bb = elem.get_BoundingBox(view)
        if bb:
            boxes.append((
                bb.Min.X - padding - eps, bb.Min.Y - padding - eps,
                bb.Max.X + padding + eps, bb.Max.Y + padding + eps,
            ))
    return boxes


def union_outlines(boxes):
    """Compute outlines of the union of axis-aligned rectangles.

    Uses coordinate-compression to build a grid, marks cells that fall inside
    any rectangle, then traces directed boundary edges (CW winding so that
    the filled region is to the right of each edge).

    Returns a list of polygons.  Each polygon is a list of (x, y) tuples
    in clockwise order with collinear vertices removed.
    """
    if not boxes:
        return []

    # Unique sorted coordinates
    xs = sorted(set(x for b in boxes for x in (b[0], b[2])))
    ys = sorted(set(y for b in boxes for y in (b[1], b[3])))

    nx, ny = len(xs) - 1, len(ys) - 1
    if nx <= 0 or ny <= 0:
        return []

    # Build occupancy grid  --  cell (i, j) spans xs[i]..xs[i+1], ys[j]..ys[j+1]
    grid = [[False] * ny for _ in range(nx)]
    for bx0, by0, bx1, by1 in boxes:
        ix0, ix1 = xs.index(bx0), xs.index(bx1)
        iy0, iy1 = ys.index(by0), ys.index(by1)
        for i in range(ix0, ix1):
            for j in range(iy0, iy1):
                grid[i][j] = True

    def filled(i, j):
        return 0 <= i < nx and 0 <= j < ny and grid[i][j]

    # Collect directed boundary edges (CW: filled region on the right)
    adj = defaultdict(list)
    for i in range(nx):
        for j in range(ny):
            if not grid[i][j]:
                continue
            if not filled(i, j - 1):          # bottom
                adj[(xs[i + 1], ys[j])].append((xs[i], ys[j]))
            if not filled(i, j + 1):          # top
                adj[(xs[i], ys[j + 1])].append((xs[i + 1], ys[j + 1]))
            if not filled(i - 1, j):          # left
                adj[(xs[i], ys[j])].append((xs[i], ys[j + 1]))
            if not filled(i + 1, j):          # right
                adj[(xs[i + 1], ys[j + 1])].append((xs[i + 1], ys[j]))

    if not adj:
        return []

    # ------------------------------------------------------------------
    # Chain edges into closed polygons.
    # At vertices with multiple outgoing edges (happens at pinch-points)
    # pick the most-clockwise turn relative to the incoming direction.
    # ------------------------------------------------------------------

    def _direction(a, b):
        """Unit-ish direction from a to b (only axis-aligned)."""
        dx = 1 if b[0] > a[0] else (-1 if b[0] < a[0] else 0)
        dy = 1 if b[1] > a[1] else (-1 if b[1] < a[1] else 0)
        return (dx, dy)

    def _cw_priority(incoming, outgoing):
        """Lower value = more clockwise turn from *incoming* direction."""
        dx, dy = incoming
        # CW preference order: right-turn, straight, left-turn, u-turn
        order = [(-dy, dx), (dx, dy), (dy, -dx), (-dx, -dy)]
        try:
            return order.index(outgoing)
        except ValueError:
            return 99

    used_edges = set()
    polygons = []

    for start_pt in list(adj.keys()):
        for first_end in list(adj[start_pt]):
            edge_key = (start_pt, first_end)
            if edge_key in used_edges:
                continue

            polygon = [start_pt]
            prev = start_pt
            curr = first_end
            used_edges.add(edge_key)

            while curr != start_pt:
                polygon.append(curr)
                inc_dir = _direction(prev, curr)
                # Pick most-CW next edge
                candidates = [e for e in adj[curr] if (curr, e) not in used_edges]
                if not candidates:
                    break
                candidates.sort(key=lambda e: _cw_priority(inc_dir, _direction(curr, e)))
                nxt = candidates[0]
                used_edges.add((curr, nxt))
                prev = curr
                curr = nxt

            # Remove collinear intermediate vertices
            merged = []
            n = len(polygon)
            for i in range(n):
                p = polygon[(i - 1) % n]
                c = polygon[i]
                q = polygon[(i + 1) % n]
                cross = (c[0] - p[0]) * (q[1] - p[1]) - (c[1] - p[1]) * (q[0] - p[0])
                if abs(cross) > 1e-10:
                    merged.append(c)

            if len(merged) >= 3:
                polygons.append(merged)

    return polygons


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

boxes = get_padded_bboxes(elements, active_view, PADDING)
if not boxes:
    forms.alert("Could not get bounding boxes for the selected elements.",
                exitscript=True)

outlines = union_outlines(boxes)
if not outlines:
    forms.alert("Could not compute outline for the selected elements.",
                exitscript=True)

cloud_count = 0
with revit.Transaction("Create Revision Cloud"):
    for outline in outlines:
        curves = List[DB.Curve]()
        for i in range(len(outline)):
            x1, y1 = outline[i]
            x2, y2 = outline[(i + 1) % len(outline)]
            p1 = DB.XYZ(x1, y1, 0.0)
            p2 = DB.XYZ(x2, y2, 0.0)
            curves.Add(DB.Line.CreateBound(p1, p2))
        DB.RevisionCloud.Create(doc, active_view, revision_id, curves)
        cloud_count += 1

print("Created {} revision cloud(s) around {} elements.".format(
    cloud_count, len(elements)))
