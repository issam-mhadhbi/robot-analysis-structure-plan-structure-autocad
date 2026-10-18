"""Robot columns -> AutoCAD sections (hatched, labelled), level by level.

Run:  python columns_to_autocad.py
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

    def add_text(self, x, y, z, label, height, layer_name):
        text = self.model.AddText(str(label), self.point(x, y, z), height)
        text.Layer = layer_name
        text.Color = self.BY_LAYER
        text.Alignment = self.ALIGN_MIDDLE_CENTER
        text.TextAlignmentPoint = self.point(x, y, z)
        return text

    def add_bubble(self, x, y, z, label, layer_name):
        circle = self.model.AddCircle(self.point(x, y, z), self.style.bubble_radius)
        circle.Layer = layer_name
        circle.Color = self.BY_LAYER
        text = self.add_text(x, y, z, label, self.style.text_height, layer_name)
        return circle, text

    def add_rectangle(self, left, bottom, right, top, layer_name):
        coords = (left, bottom, right, bottom, right, top, left, top)
        pts = VARIANT(
            pythoncom.VT_ARRAY | pythoncom.VT_R8, tuple(float(c) for c in coords)
        )
        poly = self.model.AddLightWeightPolyline(pts)
        poly.Closed = True
        poly.Layer = layer_name
        poly.Color = self.BY_LAYER
        poly.Linetype = "Continuous"
        return poly

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

    def add_polyline(self, verts, layer_name):
        """Closed continuous polyline with arcs.
        verts = [(x, y, bulge), ...]; bulge = tan(angle / 4) of the arc that
        goes from this vertex to the next one (0 = straight segment,
        positive = counter-clockwise)."""
        flat = tuple(float(c) for v in verts for c in v[:2])
        poly = self.model.AddLightWeightPolyline(
            VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, flat)
        )
        poly.Closed = True
        for i, v in enumerate(verts):
            if v[2]:
                poly.SetBulge(i, float(v[2]))
        poly.Layer = layer_name
        poly.Color = self.BY_LAYER
        poly.Linetype = "Continuous"
        return poly

    def add_circle(self, x, y, z, diameter, layer_name):
        circle = self.model.AddCircle(self.point(x, y, z), diameter / 2.0)
        circle.Layer = layer_name
        circle.Color = self.BY_LAYER
        return circle

    def add_point(self, x, y, z, layer_name):
        pt = self.model.AddPoint(self.point(x, y, z))
        pt.Layer = layer_name
        pt.Color = self.BY_LAYER
        return pt


class IRobotLabelType(IntEnum):
    I_LT_UNKNOWN = -1
    I_LT_SUPPORT = 0
    I_LT_NODE_RELEASE = 1
    I_LT_NODE_COMPATIBILITY = 2
    I_LT_BAR_SECTION = 3
    I_LT_BAR_RELEASE = 4
    I_LT_BAR_OFFSET = 5
    I_LT_BAR_CABLE = 6
    I_LT_BAREND_BRACKET = 7
    I_LT_MATERIAL = 8
    I_LT_PANEL_THICKNESS = 11
    I_LT_PANEL_REINFORCEMENT = 12
    I_LT_LINEAR_RELEASE = 13
    I_LT_BAR_ELASTIC_GROUND = 14
    I_LT_NODE_RIGID_LINK = 15
    I_LT_MEMBER_TYPE = 16
    I_LT_VEHICLE = 17
    I_LT_SOLID_PROPERTIES = 18
    I_LT_BAR_GEO_IMPERFECTIONS = 19
    I_LT_BAR_NONLINEAR_HINGE = 20
    I_LT_CLADDING = 21
    I_LT_PANEL_CALC_MODEL = 22
    I_LT_MEMBER_REINFORCEMENT_PARAMS = 25


class RobotBarSectionShapeType(IntEnum):
    I_BSST_CONCR_BEAM = -1
    I_BSST_CONCR_BEAM_T = -2
    I_BSST_CONCR_BEAM_RECT = -3
    I_BSST_CONCR_BEAM_I = -4

    I_BSST_CONCR_COL_C = -103
    I_BSST_CONCR_COL_CH = -102
    I_BSST_CONCR_COL_CQ = -101
    I_BSST_CONCR_COL_P = -104
    I_BSST_CONCR_COL_Z = -105
    I_BSST_CONCR_COL_L = -106
    I_BSST_CONCR_COL_T = -107
    I_BSST_CONCR_COL_R = -108

    I_BSST_WOOD_RECT = 41
    I_BSST_RECT_FILLED = 43
    I_BSST_CIRC_FILLED = 44
    I_BSST_CCL = 45
    I_BSST_URND = 46
    I_BSST_TRND = 47
    I_BSST_CUAP = 48
    I_BSST_SPEC_CASTELLATED_WEB_HEXAGONAL_OPENINGS = 50
    I_BSST_SPEC_CASTELLATED_WEB_ROUND_OPENINGS = 51
    I_BSST_SPEC_CASTELLATED_WEB_HEXAGONAL_OPENINGS_SHIFTED = 52
    I_BSST_SPEC_SFB = 53
    I_BSST_SPEC_IFBA = 54
    I_BSST_SPEC_IFBB = 55

    I_BSST_WOOD_DRECT = 99
    I_BSST_WOOD_CIRC = 100
    I_BSST_USER_CIRC_FILLED = 101
    I_BSST_USER_POLYGONAL = 103
    I_BSST_SPEC_CORRUGATED_WEB = 104
    I_BSST_USER_BOX_3 = 105

    I_BSST_USER_BOX = 91
    I_BSST_USER_RECT = 92
    I_BSST_USER_TUBE = 93
    I_BSST_USER_I_BISYM = 94
    I_BSST_USER_I_MONOSYM = 95
    I_BSST_USER_T_SHAPE = 96
    I_BSST_USER_C_SHAPE = 97
    I_BSST_USER_CROSS = 201

    I_BSST_COLD_C_PLUS = 1010
    I_BSST_COLD_SIGMA_SL = 1011
    I_BSST_COLD_SIGMA = 1012
    I_BSST_COLD_Z = 1013
    I_BSST_COLD_L_LIPS = 1014
    I_BSST_COLD_Z_ROT = 1015

    I_BSST_COMP_2C_FACE = 1101
    I_BSST_COMP_2C_BACK = 1102
    I_BSST_COMP_2I = 1103
    I_BSST_COMP_CI = 1104
    I_BSST_COMP_2LI = 1105
    I_BSST_COMP_4L_FACE = 1106
    I_BSST_COMP_4L_BACK = 1107
    I_BSST_COMP_2L_SHORT = 1108
    I_BSST_COMP_2L_LONG = 1109
    I_BSST_COMP_2L_CROSS = 1110
    I_BSST_COMP_2L_FACE_SHORT = 1111
    I_BSST_COMP_2L_FACE_LONG = 1112
    I_BSST_COMP_CI_BACK = 1113

    I_BSST_COMP_2C_FACE_WELD = 1201
    I_BSST_COMP_2C_BACK_WELD = 1202
    I_BSST_COMP_2I_WELD = 1203
    I_BSST_COMP_CI_WELD = 1204
    I_BSST_COMP_2LI_WELD = 1205
    I_BSST_COMP_4L_FACE_WELD = 1206
    I_BSST_COMP_4L_BACK_WELD = 1207
    I_BSST_COMP_2L_SHORT_WELD = 1208
    I_BSST_COMP_2L_LONG_WELD = 1209
    I_BSST_COMP_2L_CROSS_WELD = 1210

    I_BSST_JOIST_K = 500
    I_BSST_JOIST_LH = 501
    I_BSST_JOIST_KCS = 502
    I_BSST_JOIST_DLH = 503
    I_BSST_JOIST_SLH = 504
    I_BSST_JOIST_G = 510
    I_BSST_JOIST_VG = 511
    I_BSST_JOIST_BG = 512


class RobotBarSectionConcreteDataValue(IntEnum):
    # Column parameters
    I_BSCDV_COL_H = 0
    I_BSCDV_COL_B = 1
    I_BSCDV_COL_H1 = 2
    I_BSCDV_COL_L1 = 3
    I_BSCDV_COL_H2 = 4
    I_BSCDV_COL_L2 = 5
    I_BSCDV_COL_DE = 6
    I_BSCDV_COL_N = 7

    # Beam parameters
    I_BSCDV_BEAM_H = 0
    I_BSCDV_BEAM_B = 1
    I_BSCDV_BEAM_PL = 2
    I_BSCDV_BEAM_RL = 3
    I_BSCDV_BEAM_EL1 = 4
    I_BSCDV_BEAM_EL2 = 5
    I_BSCDV_BEAM_PR = 6
    I_BSCDV_BEAM_RR = 7
    I_BSCDV_BEAM_ER1 = 8
    I_BSCDV_BEAM_ER2 = 9
    I_BSCDV_BEAM_HG1 = 10
    I_BSCDV_BEAM_BG1 = 11
    I_BSCDV_BEAM_HG2 = 12
    I_BSCDV_BEAM_BG2 = 13

    # T-beam parameters
    I_BSCDV_BEAM_T_HF = 100
    I_BSCDV_BEAM_T_BF = 101

    # I-beam parameters
    I_BSCDV_BEAM_I_HF1 = 10
    I_BSCDV_BEAM_I_HF2 = 12
    I_BSCDV_BEAM_I_B1 = 11
    I_BSCDV_BEAM_I_B2 = 13


# ======================================================================
# ROBOT: columns
# ======================================================================
class Columns:
    def __init__(self, robot):
        self.robot = robot
        self.nodes = list()
        self.columns = list()
        self.get_all_nodes()
        self.get_all_columns()

    def get_all_nodes(self):
        nodes = self.robot.Project.Structure.Nodes
        all_nodes = nodes.GetAll() 
        for i in range(1, all_nodes.Count + 1):
            node = all_nodes.Get(i)
            node = CastTo(node, "IRobotNode")
            point = {"id": node.Number, "x": node.X, "y": node.Y, "z": node.Z}
            self.nodes.append(point)

    def get_all_columns(self):
        bars = self.robot.Project.Structure.Bars
        all_bars = bars.GetAll()
        for idx in range(1, all_bars.Count + 1):
            if True : #bars.Exist(idx) != 0:
                bar = all_bars.Get(idx)
                if bar.StructuralType == 2:
                    end_node = bar.EndNode
                    start_node = bar.StartNode

                    bar_itm = {
                        "id": bar.Number,
                        "end_xyz": [i for i in self.nodes if i["id"] == end_node][0],
                        "start_xyz": [i for i in self.nodes if i["id"] == start_node][
                            0
                        ],
                        "section": self.parseSection(bar),
                    }
                    self.columns.append(bar_itm)
                    print(json.dumps(bar_itm))

    # -- section -------------------------------------------------------------
    def parseSection(self, bar):
        SH = RobotBarSectionShapeType
        CB = RobotBarSectionConcreteDataValue

        # Parameters read for each concrete COLUMN shape
        # (l -> x direction, h -> y direction):
        #   R  : b (x) x h (y) rectangle
        #   C  : circle, diameter de
        #   CH : half circle, cut on Oy, part on x+
        #   CQ : quarter circle of radius de, in the (x+, y+) quadrant
        #   P  : regular polygon, n faces, de = 2 x apothem
        #   L  : in (x+, y+); b x h box, empty area l1 (x) x h1 (y)
        #   Z  : b x h box; empty top-left l2 x h2, empty bottom-right l1 x h1
        #   T  : b x h box; empty top-left l2 x h2, empty top-right l1 x h1
        COLUMN_PARAMS = {
            SH.I_BSST_CONCR_COL_R: {"b": CB.I_BSCDV_COL_B, "h": CB.I_BSCDV_COL_H},
            SH.I_BSST_CONCR_COL_C: {"de": CB.I_BSCDV_COL_DE},
            SH.I_BSST_CONCR_COL_CH: {"de": CB.I_BSCDV_COL_DE},
            SH.I_BSST_CONCR_COL_CQ: {"de": CB.I_BSCDV_COL_DE},
            SH.I_BSST_CONCR_COL_P: {"de": CB.I_BSCDV_COL_DE, "n": CB.I_BSCDV_COL_N},
            SH.I_BSST_CONCR_COL_L: {
                "b": CB.I_BSCDV_COL_B,
                "h": CB.I_BSCDV_COL_H,
                "h1": CB.I_BSCDV_COL_H1,
                "l1": CB.I_BSCDV_COL_L1,
            },
            SH.I_BSST_CONCR_COL_Z: {
                "b": CB.I_BSCDV_COL_B,
                "h": CB.I_BSCDV_COL_H,
                "h1": CB.I_BSCDV_COL_H1,
                "l1": CB.I_BSCDV_COL_L1,
                "h2": CB.I_BSCDV_COL_H2,
                "l2": CB.I_BSCDV_COL_L2,
            },
            SH.I_BSST_CONCR_COL_T: {
                "b": CB.I_BSCDV_COL_B,
                "h": CB.I_BSCDV_COL_H,
                "h1": CB.I_BSCDV_COL_H1,
                "l1": CB.I_BSCDV_COL_L1,
                "h2": CB.I_BSCDV_COL_H2,
                "l2": CB.I_BSCDV_COL_L2,
            },
        }

        # Robot API returns SI units (metres). Set to 100.0 to get centimetres.
        LENGTH_SCALE = 1.0
        if not bar.HasLabel(IRobotLabelType.I_LT_BAR_SECTION):
            raise ValueError(f"bar {bar.Number} has no section assigned")

        label = bar.GetLabel(IRobotLabelType.I_LT_BAR_SECTION)
        data = CastTo(label.Data, "IRobotBarSectionData")

        try:
            shape = SH(data.ShapeType)
            shape_name = shape.name
        except ValueError:
            shape = None
            shape_name = f"UNKNOWN({data.ShapeType})"

        result = {"name": label.Name, "shape": shape_name, "goe": {}}

        # rotation of the section (see ColumnsDrawer.GAMMA_UNIT / GAMMA_SIGN)
        try:
            result["goe"]["gamma"] = data.GetValue(24)
        except Exception:
            result["goe"]["gamma"] = 0.0
        result["goe"]["gamma"] = result["goe"]["gamma"] + 90
        # Concrete dimensions must be read from the Concrete sub-object
        # (data.GetValue() would return general properties: area, ...).
        params = COLUMN_PARAMS.get(shape)
        if params:
            concrete = CastTo(data.Concrete, "IRobotBarSectionConcreteData")
            for key, idx in params.items():
                value = concrete.GetValue(int(idx))
                # "n" is a count (number of polygon sides), do not scale it
                result["goe"][key] = value if key == "n" else value * LENGTH_SCALE

        return result


# ======================================================================
# AUTOCAD: column sections
# ======================================================================
class ColumnsDrawer:
    """Draws the sections of the columns (from `Columns`) in AutoCAD, level by
    level, at the position given by PointLocator.pointxyz_to_cad_xy.

    NORMAL columns
      take the column -> its HIGHEST point -> if that point belongs to a
      level (same elevation) -> draw the section in that level.

    IMPLANTED columns
      take the column -> its LOWEST point -> if that point belongs to a level
      -> look for a column whose HIGHEST point is in the same level and at
      the same position (x, y):
          found     -> a normal column already covers this position: skip
          not found -> draw the section in that level as an implanted column

    Section geometry
      Every shape is built in LOCAL coordinates, with the column node at the
      local origin, then rotated by gamma around that node and translated to
      the CAD position of the node.
        R, C, P, Z, T : centred on the node (bounding box centre / circle centre)
        CH, CQ        : circle centre on the node
        L             : corner of the b x h box on the node, box in (x+, y+)

    Style
      normal    : layer "Poteaux" (green), hatch HATCH_NORMAL
      implanted : own layers in red shades, hatch HATCH_IMPLANTED
    Each section gets a text (TEXT_PREFIX + column id) on a text layer.
    """

    # ---- normal columns ----------------------------------------------------
    SECTION_LAYER = "Poteaux"
    TEXT_LAYER = "Poteaux - Textes"
    SECTION_COLOR = 3  # green
    TEXT_COLOR = 2  # yellow
    HATCH_NORMAL = "ANSI31"  # shared by all normal columns

    # ---- implanted columns (red shades) ------------------------------------
    IMPLANTED_SECTION_LAYER = "Poteaux implantes"
    IMPLANTED_TEXT_LAYER = "Poteaux implantes - Textes"
    IMPLANTED_SECTION_COLOR = 1  # red   (outline + hatch, hatch is BYLAYER)
    IMPLANTED_TEXT_COLOR = 10  # light red shade
    HATCH_IMPLANTED = "ANSI37"  # implanted columns only

    HATCH_SCALE = 0.02  # tune if the hatch looks too dense / too sparse
    HATCH_PATTERN_PREDEFINED = 1  # acHatchPatternTypePreDefined
    HATCH_COLOR = 256  # BYLAYER

    TEXT_PREFIX = "P"
    TEXT_HEIGHT = 0.15
    TEXT_GAP = 0.10

    # ---- rotation (gamma) --------------------------------------------------
    GAMMA_UNIT = "deg"  # "deg" or "rad": unit of data.GetValue(24)
    GAMMA_SIGN = 1  # 1 = counter-clockwise in plan; -1 flips the direction

    # ---- shapes ------------------------------------------------------------
    CH_DIAMETER_IS_DE = True  # CH: de is the diameter (False: de is the radius)
    CENTERED_Z = True  # Z box centred on the node (False: corner on the node)
    CENTERED_T = True  # T box centred on the node (False: corner on the node)

    DEFAULT_SIZE = 0.30  # used when a section has no usable dimensions
    Z_TOL = 1e-3  # tolerance on elevations (m)
    XY_TOL = 0.05  # tolerance to decide "same position" (m)

    # Robot shape name -> method building the outline of that section
    SECTION_DRAWERS = {
        "I_BSST_CONCR_COL_R": "_outline_r",
        "I_BSST_CONCR_COL_C": "_outline_c",
        "I_BSST_CONCR_COL_CH": "_outline_ch",
        "I_BSST_CONCR_COL_CQ": "_outline_cq",
        "I_BSST_CONCR_COL_P": "_outline_p",
        "I_BSST_CONCR_COL_L": "_outline_l",
        "I_BSST_CONCR_COL_Z": "_outline_z",
        "I_BSST_CONCR_COL_T": "_outline_t",
    }

    def __init__(self, columns, locator):
        self.columns = list(getattr(columns, "columns", columns))
        self.locator = locator
        self.grid_drawer = locator.drawer
        self.session = self.grid_drawer.session
        self.z_cad = self.grid_drawer.COMMON_LEVEL

        self.factory = EntityFactory(self.session, self.grid_drawer.style, "Continuous")
        layers = self.grid_drawer.layers
        self.section_layer = layers.get_or_create(
            self.SECTION_LAYER, self.SECTION_COLOR
        )
        self.text_layer = layers.get_or_create(self.TEXT_LAYER, self.TEXT_COLOR)
        self.imp_section_layer = layers.get_or_create(
            self.IMPLANTED_SECTION_LAYER, self.IMPLANTED_SECTION_COLOR
        )
        self.imp_text_layer = layers.get_or_create(
            self.IMPLANTED_TEXT_LAYER, self.IMPLANTED_TEXT_COLOR
        )

        self.levels = self._levels()  # {level_name: elevation}
        self.drawn = []
        self.skipped = []  # (column id, level, reason)

    # ---- levels / column helpers -----------------------------------------
    def _levels(self):
        """{level_name: elevation}, ordered by elevation."""
        levels = {}
        for (_grid, level), e in self.locator.index.items():
            levels.setdefault(level, e["z"])
        return dict(sorted(levels.items(), key=lambda kv: kv[1]))

    @staticmethod
    def _top(col):
        a, b = col["start_xyz"], col["end_xyz"]
        return a if a["z"] >= b["z"] else b

    @staticmethod
    def _bottom(col):
        a, b = col["start_xyz"], col["end_xyz"]
        return a if a["z"] < b["z"] else b

    def _same_xy(self, p, q):
        return (
            abs(p["x"] - q["x"]) <= self.XY_TOL and abs(p["y"] - q["y"]) <= self.XY_TOL
        )

    def normal_columns(self, elevation):
        """Columns whose HIGHEST point belongs to the level at `elevation`."""
        return [
            c for c in self.columns if abs(self._top(c)["z"] - elevation) <= self.Z_TOL
        ]

    def implanted_columns(self, elevation, normal):
        """Columns whose LOWEST point belongs to the level at `elevation`
        and that have NO normal column of this level at the same position."""
        implanted = []
        for c in self.columns:
            bottom = self._bottom(c)
            if abs(bottom["z"] - elevation) > self.Z_TOL:
                continue
            covered = any(
                self._same_xy(self._top(n), bottom) for n in normal if n is not c
            )
            if not covered:
                implanted.append(c)
        return implanted

    def classify(self, elevation):
        """(normal, implanted) columns of the level at `elevation`."""
        normal = self.normal_columns(elevation)
        return normal, self.implanted_columns(elevation, normal)

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

    # ---- outline helpers ---------------------------------------------------
    # An outline is ("circle", diameter)  or  ("poly", [(x, y, bulge), ...])
    # in local coordinates (column node = origin, x to the right, y up).
    @staticmethod
    def _poly(pts, dx=0.0, dy=0.0):
        """Straight-segment outline from [(x, y), ...], shifted by (dx, dy);
        consecutive duplicate points (zero-size notches) are removed."""
        clean = []
        for x, y in pts:
            p = (x + dx, y + dy)
            if not clean or (
                abs(p[0] - clean[-1][0]) > 1e-9 or abs(p[1] - clean[-1][1]) > 1e-9
            ):
                clean.append(p)
        if len(clean) > 1 and (
            abs(clean[0][0] - clean[-1][0]) <= 1e-9
            and abs(clean[0][1] - clean[-1][1]) <= 1e-9
        ):
            clean.pop()
        return ("poly", [(x, y, 0.0) for x, y in clean])

    def _rect(self, w, h):  # rectangle centred on the node
        return self._poly(
            [
                (-w / 2.0, -h / 2.0),
                (w / 2.0, -h / 2.0),
                (w / 2.0, h / 2.0),
                (-w / 2.0, h / 2.0),
            ]
        )

    def _fallback(self, goe):
        """Used when the dimensions of a section are missing / inconsistent."""
        b, h, de = goe.get("b", 0), goe.get("h", 0), goe.get("de", 0)
        if b > 0 and h > 0:
            return self._rect(b, h)
        if de > 0:
            return self._rect(de, de)
        return self._rect(self.DEFAULT_SIZE, self.DEFAULT_SIZE)

    @staticmethod
    def _has_dims(goe):
        return any(goe.get(k, 0) > 0 for k in ("b", "h", "de"))

    # ---- one outline function per column section type ---------------------
    def _outline_r(self, goe):
        """Rectangle: b in x, h in y, centred."""
        b, h = goe.get("b", 0), goe.get("h", 0)
        if b <= 0 or h <= 0:
            return self._fallback(goe)
        return self._rect(b, h)

    def _outline_c(self, goe):
        """Circle of diameter de, centred."""
        de = goe.get("de", 0)
        if de <= 0:
            return self._fallback(goe)
        return ("circle", de)

    def _outline_ch(self, goe):
        """Half circle, cut on the Oy axis, the part is on the x+ side."""
        de = goe.get("de", 0)
        if de <= 0:
            return self._fallback(goe)
        r = de / 2.0 if self.CH_DIAMETER_IS_DE else de
        # arc from (0,-r) to (0,r) counter-clockwise (through x+), then back
        return ("poly", [(0.0, -r, 1.0), (0.0, r, 0.0)])

    def _outline_cq(self, goe):
        """Quarter of a circle of radius de, in the (x+, y+) quadrant."""
        r = goe.get("de", 0)
        if r <= 0:
            return self._fallback(goe)
        return (
            "poly",
            [(0.0, 0.0, 0.0), (r, 0.0, math.tan(math.pi / 8.0)), (0.0, r, 0.0)],
        )

    def _outline_p(self, goe):
        """Regular polygon with n faces, de = 2 x apothem, centred."""
        de, n = goe.get("de", 0), int(round(goe.get("n", 0) or 0))
        if de <= 0 or n < 3:
            return self._fallback(goe)
        R = (de / 2.0) / math.cos(math.pi / n)  # circumscribed radius
        pts = [
            (
                R * math.cos(math.pi / 2.0 + math.pi / n + 2.0 * math.pi * i / n),
                R * math.sin(math.pi / 2.0 + math.pi / n + 2.0 * math.pi * i / n),
            )
            for i in range(n)
        ]  # flat faces at the top and bottom
        return self._poly(pts)

    def _outline_l(self, goe):
        """L in the (x+, y+) quadrant: b x h box, empty area l1 (x) x h1 (y)
        in its top-right corner."""
        b, h = goe.get("b", 0), goe.get("h", 0)
        h1, l1 = goe.get("h1", 0), goe.get("l1", 0)
        if b <= 0 or h <= 0:
            return self._fallback(goe)
        if not (0 <= l1 < b and 0 <= h1 < h):
            return self._poly([(0, 0), (b, 0), (b, h), (0, h)])
        return self._poly(
            [(0, 0), (b, 0), (b, h - h1), (b - l1, h - h1), (b - l1, h), (0, h)]
        )

    def _outline_z(self, goe):
        """Z: b x h box, empty top-left (l2 x h2) and bottom-right (l1 x h1)."""
        b, h = goe.get("b", 0), goe.get("h", 0)
        h1, l1 = goe.get("h1", 0), goe.get("l1", 0)
        h2, l2 = goe.get("h2", 0), goe.get("l2", 0)
        if b <= 0 or h <= 0:
            return self._fallback(goe)
        dx, dy = (-b / 2.0, -h / 2.0) if self.CENTERED_Z else (0.0, 0.0)
        if not (0 <= l1 < b and 0 <= l2 < b and 0 <= h1 < h and 0 <= h2 < h):
            return self._poly([(0, 0), (b, 0), (b, h), (0, h)], dx, dy)
        pts = [
            (0, 0),
            (b - l1, 0),
            (b - l1, h1),
            (b, h1),
            (b, h),
            (l2, h),
            (l2, h - h2),
            (0, h - h2),
        ]
        return self._poly(pts, dx, dy)

    def _outline_t(self, goe):
        """T: b x h box, empty top-left (l2 x h2) and top-right (l1 x h1)."""
        b, h = goe.get("b", 0), goe.get("h", 0)
        h1, l1 = goe.get("h1", 0), goe.get("l1", 0)
        h2, l2 = goe.get("h2", 0), goe.get("l2", 0)
        if b <= 0 or h <= 0:
            return self._fallback(goe)
        dx, dy = (-b / 2.0, -h / 2.0) if self.CENTERED_T else (0.0, 0.0)
        if not (l1 >= 0 and l2 >= 0 and l1 + l2 < b and 0 <= h1 < h and 0 <= h2 < h):
            return self._poly([(0, 0), (b, 0), (b, h), (0, h)], dx, dy)
        pts = [
            (0, 0),
            (b, 0),
            (b, h - h1),
            (b - l1, h - h1),
            (b - l1, h),
            (l2, h),
            (l2, h - h2),
            (0, h - h2),
        ]
        return self._poly(pts, dx, dy)

    def _outline_default(self, goe):  # unknown shape
        return self._fallback(goe)

    # ---- rotation / extents ------------------------------------------------
    def _gamma_rad(self, goe):
        g = float(goe.get("gamma", 0.0) or 0.0)
        if self.GAMMA_UNIT == "deg":
            g = math.radians(g)
        return g * self.GAMMA_SIGN

    @staticmethod
    def _rotate(verts, angle):
        c, s = math.cos(angle), math.sin(angle)
        return [(x * c - y * s, x * s + y * c, b) for x, y, b in verts]

    @staticmethod
    def _arc_samples(p, q, bulge, n=8):
        """Points along the arc p -> q defined by `bulge` (for extents)."""
        theta = 4.0 * math.atan(bulge)
        dx, dy = q[0] - p[0], q[1] - p[1]
        chord = math.hypot(dx, dy)
        if chord == 0:
            return []
        r = chord / (2.0 * math.sin(abs(theta) / 2.0))
        d = r * math.cos(theta / 2.0)
        cx = (p[0] + q[0]) / 2.0 - dy / chord * d
        cy = (p[1] + q[1]) / 2.0 + dx / chord * d
        a0 = math.atan2(p[1] - cy, p[0] - cx)
        return [
            (
                cx + r * math.cos(a0 + theta * k / n),
                cy + r * math.sin(a0 + theta * k / n),
            )
            for k in range(1, n)
        ]

    def _top_y(self, verts):
        """Highest local y of a polyline outline (arcs included)."""
        ys = [v[1] for v in verts]
        for i, v in enumerate(verts):
            if v[2]:
                nxt = verts[(i + 1) % len(verts)]
                ys += [p[1] for p in self._arc_samples(v[:2], nxt[:2], v[2])]
        return max(ys)

    # ---- hatch -----------------------------------------------------------
    def _add_hatch(self, boundary, pattern, layer):
        try:
            hatch = self.session.model.AddHatch(
                self.HATCH_PATTERN_PREDEFINED, pattern, True
            )
            loop = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, tuple(boundary))
            hatch.AppendOuterLoop(loop)
            hatch.PatternScale = self.HATCH_SCALE
            hatch.Layer = layer
            hatch.Color = self.HATCH_COLOR  # BYLAYER -> green / red
            hatch.Evaluate()
            return hatch
        except Exception as exc:
            print("Could not hatch section ({}): {}".format(pattern, exc))
            return None

    # ---- one column ------------------------------------------------------
    def _draw_column(self, col, level_name, elevation, implanted):
        # normal: position of the top point | implanted: position of the bottom
        node = self._bottom(col) if implanted else self._top(col)
        pos = self._cad_xy(node["x"], node["y"], level_name, elevation)
        if pos is None:
            self.skipped.append((col["id"], level_name, "no CAD position"))
            return
        cx, cy = pos

        section = col["section"]
        goe = section.get("goe", {})
        if not self._has_dims(goe):
            print(
                "  Column {}: no dimensions for shape {} -> default {} m square".format(
                    col["id"], section.get("shape"), self.DEFAULT_SIZE
                )
            )

        if implanted:
            section_layer, text_layer = self.imp_section_layer, self.imp_text_layer
            pattern = self.HATCH_IMPLANTED
        else:
            section_layer, text_layer = self.section_layer, self.text_layer
            pattern = self.HATCH_NORMAL

        build = getattr(
            self,
            self.SECTION_DRAWERS.get(section.get("shape"), "_outline_default"),
        )
        kind, outline = build(goe)

        if kind == "circle":
            boundary = [
                self.factory.add_circle(cx, cy, self.z_cad, outline, section_layer)
            ]
            top = cy + outline / 2.0
        else:
            verts = self._rotate(outline, self._gamma_rad(goe))
            top = cy + self._top_y(verts)
            placed = [(cx + x, cy + y, b) for x, y, b in verts]
            boundary = [self.factory.add_polyline(placed, section_layer)]

        self._add_hatch(boundary, pattern, section_layer)

        self.factory.add_text(
            cx,
            top + self.TEXT_GAP + self.TEXT_HEIGHT / 2.0,
            self.z_cad,
            "{}{}".format(self.TEXT_PREFIX, col["id"]),
            self.TEXT_HEIGHT,
            text_layer,
        )

        self.drawn.append(
            {
                "id": col["id"],
                "level": level_name,
                "implanted": implanted,
                "cad_x": cx,
                "cad_y": cy,
                "section": section.get("name"),
            }
        )

    # ---- public API ------------------------------------------------------
    def draw(self):
        for level_name, elevation in self.levels.items():
            normal, implanted = self.classify(elevation)
            print(
                "\nLevel '{}' (Z={}): {} normal, {} implanted column(s)".format(
                    level_name, elevation, len(normal), len(implanted)
                )
            )
            for col in normal:
                self._draw_column(col, level_name, elevation, implanted=False)
            for col in implanted:
                self._draw_column(col, level_name, elevation, implanted=True)

        drawn_ids = {d["id"] for d in self.drawn}
        never = [c["id"] for c in self.columns if c["id"] not in drawn_ids]
        if never:
            print(
                "\nColumns not drawn (no level at their top or bottom): {}".format(
                    never
                )
            )
        for cid, level, reason in self.skipped:
            print("Skipped column {} on {}: {}".format(cid, level, reason))

        self.session.refresh_view()
        return self


def ColumnsPipeline(robot, locator):
    columns = Columns(robot)
    return ColumnsDrawer(columns, locator).draw()


if __name__ == "__main__":
    robot = win32com.client.Dispatch("Robot.Application")
    robot.Visible = True

