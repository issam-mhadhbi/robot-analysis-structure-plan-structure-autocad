"""Robot slabs (panels) -> AutoCAD, level by level.

Each slab is drawn as:
  * its contour, offset toward the inside
  * lines between opposite points of the contour (marks the slab as an area)
  * at the centre, a "fraction":   {name}
                                   ------
                                   {thickness}-{model}
  * the load direction:
        "xy"  -> a "+" with arrows at its four ends
        "x"   -> a horizontal line with arrows, under the text
        "y"   -> a vertical line with arrows, beside the text

Run:  python slabs_to_autocad.py
Needs the grid module (GridPipeline) in the same folder; see the import below.
"""

import json
import math
from enum import IntEnum

import pythoncom
import win32com.client
from win32com.client import VARIANT, CastTo




class EntityFactory:
    BY_LAYER = 256
    ALIGN_MIDDLE_CENTER = 10  # acAlignmentMiddleCenter

    def __init__(self, session, style, line_type):
        self.model = session.model
        self.style = style
        self.line_type = line_type

    @staticmethod
    def point(x, y, z=0.0):
        return VARIANT(
            pythoncom.VT_ARRAY | pythoncom.VT_R8,
            (float(x), float(y), float(z)),
        )

    def add_line(self, p1, p2, layer_name):
        line = self.model.AddLine(self.point(*p1), self.point(*p2))
        line.Layer = layer_name
        line.Color = self.BY_LAYER
        line.Linetype = self.line_type
        line.LinetypeScale = 1.0
        return line

    def add_text(self, x, y, z, label, height, layer_name, rotation=0.0):
        """Text centred on (x, y); rotation in radians (counter-clockwise)."""
        text = self.model.AddText(str(label), self.point(x, y, z), height)
        text.Layer = layer_name
        text.Color = self.BY_LAYER
        text.Alignment = self.ALIGN_MIDDLE_CENTER
        text.TextAlignmentPoint = self.point(x, y, z)
        if rotation:
            text.Rotation = float(rotation)
        return text

    def add_polygon(self, pts, layer_name):
        """Closed continuous polyline through pts [(x, y), ...]."""
        flat = tuple(float(c) for p in pts for c in p)
        poly = self.model.AddLightWeightPolyline(
            VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, flat)
        )
        poly.Closed = True
        poly.Layer = layer_name
        poly.Color = self.BY_LAYER
        poly.Linetype = "Continuous"
        return poly


class IRobotLabelType(IntEnum):
    I_LT_UNKNOWN = -1  # Unknown label type. Available since version 1.7.
    I_LT_SUPPORT = 0  # Label describing a support. Available since version 2.0.
    I_LT_NODE_RELEASE = 1  # Label describing a node release.
    I_LT_NODE_COMPATIBILITY = 2  # Label describing a compatibility.
    I_LT_BAR_SECTION = 3  # Label describing a bar section.
    I_LT_BAR_RELEASE = 4  # Label describing a bar release.
    I_LT_BAR_OFFSET = 5  # Label describing a bar offset.
    I_LT_BAR_CABLE = 6  # Label describing a cable.
    I_LT_BAREND_BRACKET = 7  # Label describing a bracket at one of bar ends.
    I_LT_MATERIAL = 8  # Label describing a material. Available since version 2.0.
    I_LT_PANEL_THICKNESS = 11  # Label describing a plate thickness.
    I_LT_PANEL_REINFORCEMENT = 12  # Label describing a plate reinforcement.
    I_LT_LINEAR_RELEASE = 13  # Linear release between panels. Available since version 2.5.
    I_LT_BAR_ELASTIC_GROUND = 14  # Elastic ground (bars). Available since version 2.5.
    I_LT_NODE_RIGID_LINK = 15  # Rigid link. Available since version 2.5.
    I_LT_MEMBER_TYPE = 16  # Member type for steel/aluminum and timber design. Available since version 2.5.
    I_LT_VEHICLE = 17  # Vehicle definition. Available since version 3.
    I_LT_SOLID_PROPERTIES = 18  # Solid properties. Available since version 3.
    I_LT_BAR_GEO_IMPERFECTIONS = 19  # Geometrical imperfections. Available since version 3.
    I_LT_BAR_NONLINEAR_HINGE = 20  # Non-linear hinge. Available since version 3.
    I_LT_CLADDING = 21  # Cladding. Available since version 3.5.
    I_LT_PANEL_CALC_MODEL = 22  # Label describing a panel calculation model. Available since version 9.7.
    I_LT_MEMBER_REINFORCEMENT_PARAMS = 25  # Member reinforcement parameters. Available since version 12.


class IRobotPanelCalcModelLoadTransfer(IntEnum):
    I_PCMLT_ANALYTICAL_FINITE_ELEMENTS = 1  # Available since version 9.7.
    I_PCMLT_SIMPLIFIED_TWO_WAY = 2  # Available since version 9.7.
    I_PCMLT_SIMPLIFIED_ONE_WAY = 3


# ======================================================================
# ROBOT: slabs
# ======================================================================
class Slabs:
    OBJECT_TYPE_PANEL = 4  # StructuralType of a slab object
    LOAD_TRANSFER = {
        IRobotPanelCalcModelLoadTransfer.I_PCMLT_ANALYTICAL_FINITE_ELEMENTS: "",
        IRobotPanelCalcModelLoadTransfer.I_PCMLT_SIMPLIFIED_TWO_WAY: "two-way",
        IRobotPanelCalcModelLoadTransfer.I_PCMLT_SIMPLIFIED_ONE_WAY: "one-way",
    }
    TOL = 1e-6  # tolerance to decide that an edge is parallel to x or y (m)

    def __init__(self, robot):
        self.robot = robot
        self.slabs = list()
        self.get_all_slabs()

    @staticmethod
    def _label(obj, label_type):
        try:
            return obj.GetLabel(label_type)
        except Exception:
            return None

    def get_all_slabs(self):
        structure = self.robot.Project.Structure
        objs = structure.Objects.GetAll()
        for i in range(1, objs.Count + 1):  # 1-based
            obj = objs.Get(i)
            if obj.StructuralType != self.OBJECT_TYPE_PANEL:
                continue

            thickness = self._label(obj, IRobotLabelType.I_LT_PANEL_THICKNESS)
            model = self._label(obj, IRobotLabelType.I_LT_PANEL_CALC_MODEL)

            load_transfer = None
            if model is not None:
                try:
                    load_transfer = self.LOAD_TRANSFER.get(model.Data.LoadTransfer)
                except Exception:
                    load_transfer = None

            slab = {
                "id": obj.Number,
                "name": obj.Name,
                "thickness": thickness.Name if thickness is not None else "",
                "model": model.Name if model is not None else "",
                "loadTransfer": load_transfer,
                "load-direction": "xy",
                "contour": list(),
            }

            points = obj.Main.DefPoints
            for k in range(1, points.Count + 1):
                p = points.Get(k)
                slab["contour"].append({"x": p.X, "y": p.Y, "z": p.Z})

            if slab["loadTransfer"] == "one-way" and len(slab["contour"]) >= 2:
                p1, p2 = slab["contour"][0], slab["contour"][1]
                if abs(p1["x"] - p2["x"]) <= self.TOL:
                    slab["load-direction"] = "y"
                elif abs(p1["y"] - p2["y"]) <= self.TOL:
                    slab["load-direction"] = "x"
                else:
                    print(
                        "Slab {}: first edge is neither parallel to x nor y, "
                        "load direction set to xy.".format(slab["id"])
                    )

            self.slabs.append(slab)
            print(json.dumps(slab))
        print("{} slab(s) read from Robot.".format(len(self.slabs)))


# ======================================================================
# AUTOCAD: slabs
# ======================================================================
class SlabsDrawer:
    """Draws the slabs (from `Slabs`) in AutoCAD, level by level, at the
    position given by PointLocator.pointxyz_to_cad_xy.

    Rule for a level L (elevation zL): a slab is drawn in L when ALL the
    points of its contour are at zL (inclined slabs are listed, not drawn).
    """

    CONTOUR_LAYER = "Dalles"
    MARK_LAYER = "Dalles - Aire"
    TEXT_LAYER = "Dalles - Textes"
    ARROW_LAYER = "Dalles - Sens"
    CONTOUR_COLOR = 6  # magenta
    MARK_COLOR = 8  # grey
    TEXT_COLOR = 7  # white / black
    ARROW_COLOR = 30  # orange

    OFFSET = 0.15  # contour offset toward the inside (m); 0 = no offset

    # fields: {name}, {thickness}, {model}, {id}
    TEXT_TOP = "{name}"
    TEXT_BOTTOM = "{thickness}-{model}"
    TEXT_HEIGHT = 0.20
    TEXT_GAP = 0.05  # gap between the fraction bar and each text (m)
    CHAR_WIDTH = 0.65  # average character width / text height (bar length)

    ARROW_HALF = 0.50  # half length of the direction symbol (m)
    ARROW_HEAD = 0.12  # arrow head length (m)
    ARROW_ANGLE = 25.0  # arrow head half angle (degrees)

    Z_TOL = 1e-3  # tolerance on elevations (m)
    XY_TOL = 1e-6

    def __init__(self, slabs, locator):
        self.slabs = list(getattr(slabs, "slabs", slabs))
        self.locator = locator
        self.grid_drawer = locator.drawer
        self.session = self.grid_drawer.session
        self.z_cad = self.grid_drawer.COMMON_LEVEL

        self.factory = EntityFactory(self.session, self.grid_drawer.style, "Continuous")
        layers = self.grid_drawer.layers
        self.contour_layer = layers.get_or_create(
            self.CONTOUR_LAYER, self.CONTOUR_COLOR
        )
        self.mark_layer = layers.get_or_create(self.MARK_LAYER, self.MARK_COLOR)
        self.text_layer = layers.get_or_create(self.TEXT_LAYER, self.TEXT_COLOR)
        self.arrow_layer = layers.get_or_create(self.ARROW_LAYER, self.ARROW_COLOR)

        self.levels = self._levels()  # {level_name: elevation}
        self.drawn = []
        self.skipped = []  # (slab id, level, reason)

    # ---- levels / helpers ------------------------------------------------
    def _levels(self):
        levels = {}
        for (_grid, level), e in self.locator.index.items():
            levels.setdefault(level, e["z"])
        return dict(sorted(levels.items(), key=lambda kv: kv[1]))

    def _flat(self, slab):
        zs = [p["z"] for p in slab["contour"]]
        return bool(zs) and max(zs) - min(zs) <= self.Z_TOL

    def slabs_of_level(self, elevation):
        return [
            s
            for s in self.slabs
            if self._flat(s) and abs(s["contour"][0]["z"] - elevation) <= self.Z_TOL
        ]

    def _cad_xy(self, x, y, level_name, elevation):
        results = [
            r
            for r in self.locator.pointxyz_to_cad_xy(x, y, elevation, tol=self.Z_TOL)
            if r["level"] == level_name
        ]
        if not results:
            return None
        inside = [r for r in results if r["inside_grid"]]
        r = (inside or results)[0]
        return r["cad_x"], r["cad_y"]

    # ---- polygon geometry ------------------------------------------------
    def _clean(self, pts):
        """Remove consecutive duplicate points (and a repeated closing point)."""
        out = []
        for p in pts:
            if not out or math.hypot(p[0] - out[-1][0], p[1] - out[-1][1]) > self.XY_TOL:
                out.append(p)
        if len(out) > 1 and math.hypot(out[0][0] - out[-1][0], out[0][1] - out[-1][1]) <= self.XY_TOL:
            out.pop()
        return out

    @staticmethod
    def _area(pts):
        return 0.5 * sum(
            x1 * y2 - x2 * y1 for (x1, y1), (x2, y2) in zip(pts, pts[1:] + pts[:1])
        )

    def _inset(self, pts, d):
        """Polygon moved toward its inside by `d` (mitre joints).
        Returns the original polygon if the result is not valid."""
        n = len(pts)
        if d <= 0 or n < 3:
            return pts
        area = self._area(pts)
        if abs(area) < 1e-12:
            return pts
        sgn = 1.0 if area > 0 else -1.0

        edges = []  # (base point shifted inward, direction)
        for i in range(n):
            x1, y1 = pts[i]
            x2, y2 = pts[(i + 1) % n]
            dx, dy = x2 - x1, y2 - y1
            length = math.hypot(dx, dy)
            nx, ny = sgn * (-dy / length), sgn * (dx / length)  # inward normal
            edges.append(((x1 + nx * d, y1 + ny * d), (dx, dy), length))

        new = []
        for i in range(n):
            (b0, e0, l0), (b1, e1, l1) = edges[i - 1], edges[i]
            cross = e0[0] * e1[1] - e0[1] * e1[0]
            if abs(cross) < 1e-9 * l0 * l1:  # parallel edges
                new.append(b1)
                continue
            t = ((b1[0] - b0[0]) * e1[1] - (b1[1] - b0[1]) * e1[0]) / cross
            new.append((b0[0] + t * e0[0], b0[1] + t * e0[1]))

        new_area = self._area(new)
        # every offset edge must keep the direction of its original edge,
        # otherwise the offset is larger than the slab allows
        flipped = any(
            (new[(i + 1) % n][0] - new[i][0]) * edges[i][1][0]
            + (new[(i + 1) % n][1] - new[i][1]) * edges[i][1][1]
            <= 0
            for i in range(n)
        )
        if flipped or new_area * area <= 0 or abs(new_area) >= abs(area):
            print("  offset {} m not possible for this contour: not applied".format(d))
            return pts
        return new

    @staticmethod
    def _inside(pt, poly):
        x, y = pt
        res = False
        for (x1, y1), (x2, y2) in zip(poly, poly[1:] + poly[:1]):
            if (y1 > y) != (y2 > y) and x < (x2 - x1) * (y - y1) / (y2 - y1) + x1:
                res = not res
        return res

    def _centre(self, poly):
        """Area centroid; if it falls outside (concave slab) the middle of the
        bounding box, then the mean of the vertices."""
        a = self._area(poly)
        cx = cy = 0.0
        for (x1, y1), (x2, y2) in zip(poly, poly[1:] + poly[:1]):
            f = x1 * y2 - x2 * y1
            cx += (x1 + x2) * f
            cy += (y1 + y2) * f
        if abs(a) > 1e-12:
            cand = (cx / (6.0 * a), cy / (6.0 * a))
            if self._inside(cand, poly):
                return cand
        xs, ys = [p[0] for p in poly], [p[1] for p in poly]
        cand = ((min(xs) + max(xs)) / 2.0, (min(ys) + max(ys)) / 2.0)
        if self._inside(cand, poly):
            return cand
        return (sum(xs) / len(xs), sum(ys) / len(ys))

    # ---- marks -----------------------------------------------------------
    def _area_lines(self, poly):
        """Lines between opposite points of the contour."""
        n = len(poly)
        z = self.z_cad
        if n == 3:  # no opposite points: join each vertex to the centroid
            c = self._centre(poly)
            for p in poly:
                self.factory.add_line((p[0], p[1], z), (c[0], c[1], z), self.mark_layer)
            return
        step = n // 2
        pairs = set()
        for i in range(n):
            j = (i + step) % n
            pairs.add((min(i, j), max(i, j)))
        for i, j in sorted(pairs):
            self.factory.add_line(
                (poly[i][0], poly[i][1], z), (poly[j][0], poly[j][1], z), self.mark_layer
            )

    def _arrow_head(self, tip, ux, uy):
        """Two short lines forming an arrow head at `tip`, pointing along (ux, uy)."""
        a = math.radians(self.ARROW_ANGLE)
        for s in (1.0, -1.0):
            ca, sa = math.cos(a), math.sin(s * a)
            bx = -(ux * ca - uy * sa) * self.ARROW_HEAD
            by = -(ux * sa + uy * ca) * self.ARROW_HEAD
            self.factory.add_line(
                (tip[0], tip[1], self.z_cad),
                (tip[0] + bx, tip[1] + by, self.z_cad),
                self.arrow_layer,
            )

    def _arrow_line(self, p1, p2):
        """Line p1 -> p2 with an arrow head at both ends."""
        z = self.z_cad
        self.factory.add_line((p1[0], p1[1], z), (p2[0], p2[1], z), self.arrow_layer)
        dx, dy = p2[0] - p1[0], p2[1] - p1[1]
        length = math.hypot(dx, dy)
        if length == 0:
            return
        ux, uy = dx / length, dy / length
        self._arrow_head(p2, ux, uy)
        self._arrow_head(p1, -ux, -uy)

    def _direction_symbol(self, direction, cx, cy, text_w, text_bottom, half):
        margin = 0.5 * self.TEXT_HEIGHT
        if direction == "y":  # vertical line beside the text
            x = cx + text_w / 2.0 + margin + 0.1
            self._arrow_line((x, cy - half), (x, cy + half))
        elif direction == "x":  # horizontal line under the text
            y = text_bottom - margin - 0.1
            self._arrow_line((cx - half, y), (cx + half, y))
        else:  # "xy": + with arrows, under the text
            yc = text_bottom - margin - half
            self._arrow_line((cx - half, yc), (cx + half, yc))
            self._arrow_line((cx, yc - half), (cx, yc + half))

    # ---- one slab --------------------------------------------------------
    def _draw_slab(self, slab, level_name, elevation):
        pts = []
        for p in slab["contour"]:
            q = self._cad_xy(p["x"], p["y"], level_name, elevation)
            if q is None:
                self.skipped.append((slab["id"], level_name, "no CAD position"))
                return
            pts.append(q)
        pts = self._clean(pts)
        if len(pts) < 3:
            self.skipped.append((slab["id"], level_name, "contour has < 3 points"))
            return

        poly = self._inset(pts, self.OFFSET)

        # 1) contour
        self.factory.add_polygon(poly, self.contour_layer)

        # 2) lines between opposite points
        self._area_lines(poly)

        # 3) text as a fraction:  name / (thickness-model)
        fields = dict(
            id=slab["id"],
            name=slab["name"],
            thickness=slab["thickness"],
            model=slab["model"],
        )
        top = self.TEXT_TOP.format(**fields)
        bottom = self.TEXT_BOTTOM.format(**fields)
        h, gap = self.TEXT_HEIGHT, self.TEXT_GAP
        cx, cy = self._centre(poly)
        z = self.z_cad

        text_w = max(len(top), len(bottom)) * h * self.CHAR_WIDTH + 0.4 * h
        self.factory.add_text(cx, cy + gap + h / 2.0, z, top, h, self.text_layer)
        self.factory.add_text(cx, cy - gap - h / 2.0, z, bottom, h, self.text_layer)
        self.factory.add_line(
            (cx - text_w / 2.0, cy, z), (cx + text_w / 2.0, cy, z), self.text_layer
        )

        # 4) load direction
        xs, ys = [p[0] for p in poly], [p[1] for p in poly]
        half = min(self.ARROW_HALF, 0.3 * min(max(xs) - min(xs), max(ys) - min(ys)))
        self._direction_symbol(
            slab["load-direction"], cx, cy, text_w, cy - gap - h, half
        )

        self.drawn.append(
            {
                "id": slab["id"],
                "level": level_name,
                "name": slab["name"],
                "direction": slab["load-direction"],
            }
        )

    # ---- public API ------------------------------------------------------
    def draw(self):
        for level_name, elevation in self.levels.items():
            slabs = self.slabs_of_level(elevation)
            print(
                "\nLevel '{}' (Z={}): {} slab(s)".format(
                    level_name, elevation, len(slabs)
                )
            )
            for slab in slabs:
                self._draw_slab(slab, level_name, elevation)

        drawn_ids = {d["id"] for d in self.drawn}
        inclined = [s["id"] for s in self.slabs if not self._flat(s)]
        other = [s["id"] for s in self.slabs if self._flat(s) and s["id"] not in drawn_ids]
        if inclined:
            print("\nInclined slabs (not drawn in plan): {}".format(inclined))
        if other:
            print("\nSlabs not drawn (no level at their elevation): {}".format(other))
        for sid, level, reason in self.skipped:
            print("Skipped slab {} on {}: {}".format(sid, level, reason))

        self.session.refresh_view()
        return self


def SlabsPipeline(robot, locator):
    slabs = Slabs(robot)
    return SlabsDrawer(slabs, locator).draw()


if __name__ == "__main__":
    robot = win32com.client.Dispatch("Robot.Application")
    robot.Visible = True

    # 1) grids: read from Robot and drawn in a new AutoCAD document
    locator = GridPipeline(robot)

    # 2) slabs on the grids
    drawer = SlabsPipeline(robot, locator)

    print(
        "\nDone: {} slab(s) drawn, {} skipped.".format(
            len(drawer.drawn), len(drawer.skipped)
        )
    )