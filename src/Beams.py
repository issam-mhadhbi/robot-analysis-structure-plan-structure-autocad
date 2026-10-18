"""Robot beams -> AutoCAD (outline in plan + name inside, along the beam).

Run:  python beams_to_autocad.py
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
    I_LT_BAR_SECTION = 3  # Label describing a bar section.


class RobotBarSectionShapeType(IntEnum):
    I_BSST_CONCR_BEAM = -1  # RC beam of any shape.
    I_BSST_CONCR_BEAM_T = -2  # RC T-beam.
    I_BSST_CONCR_BEAM_RECT = -3  # RC rectangular beam.
    I_BSST_CONCR_BEAM_I = -4  # RC I-beam.


class RobotBarSectionConcreteDataValue(IntEnum):
    I_BSCDV_BEAM_H = 0  # Beam height.
    I_BSCDV_BEAM_B = 1  # Beam width.
    I_BSCDV_BEAM_T_HF = 100  # T-beam flange height.
    I_BSCDV_BEAM_T_BF = 101  # T-beam flange width.


# ======================================================================
# ROBOT: beams
# ======================================================================
class Beams:
    BAR_STRUCTURAL_TYPE_BEAM = 1

    def __init__(self, robot):
        self.robot = robot
        self.nodes = dict()  # {node id: {"id", "x", "y", "z"}}
        self.beams = list()
        self.get_all_nodes()
        self.get_all_beams()

    def get_all_nodes(self):
        nodes = self.robot.Project.Structure.Nodes
        all_nodes = nodes.GetAll()
        for i in range(1, all_nodes.Count + 1):
            node = CastTo(all_nodes.Get(i), "IRobotNode")
            self.nodes[node.Number] = {
                "id": node.Number,
                "x": node.X,
                "y": node.Y,
                "z": node.Z,
            }

    def get_all_beams(self):
        bars = self.robot.Project.Structure.Bars
        all_bars = bars.GetAll()
        for idx in range(1, all_bars.Count + 1):
            bar = all_bars.Get(idx)
            if bar.StructuralType != self.BAR_STRUCTURAL_TYPE_BEAM:
                continue
            if bar.StartNode not in self.nodes or bar.EndNode not in self.nodes:
                print("Beam {}: node not found, skipped.".format(bar.Number))
                continue
            try:
                section = self.parseSection(bar)
            except Exception as exc:
                print("Beam {}: {} -> default section".format(bar.Number, exc))
                raise ValueError("Beam {}: {} -> default section".format(bar.Number, exc))
                section = {"name": "?", "shape": "UNKNOWN", "goe": {"gamma": 0.0}}

            item = {
                "id": bar.Number,
                "start_xyz": self.nodes[bar.StartNode],
                "end_xyz": self.nodes[bar.EndNode],
                "section": section,
            }
            self.beams.append(item)
            print(json.dumps(item))
        print("{} beam(s) read from Robot.".format(len(self.beams)))

    # -- section -------------------------------------------------------------
    def parseSection(self, bar):
        SH = RobotBarSectionShapeType
        CB = RobotBarSectionConcreteDataValue

        # b = width, h = height of the section (T beam: b = web width,
        # hf / bf = flange height / width). In plan the beam is drawn with
        # its width b (or h when the section is rotated by 90 degrees).
        BEAM_PARAMS = {
            SH.I_BSST_CONCR_BEAM: {"b": CB.I_BSCDV_BEAM_B, "h": CB.I_BSCDV_BEAM_H},
            SH.I_BSST_CONCR_BEAM_RECT: {
                "b": CB.I_BSCDV_BEAM_B,
                "h": CB.I_BSCDV_BEAM_H,
            },
            SH.I_BSST_CONCR_BEAM_T: {
                "b": CB.I_BSCDV_BEAM_B,
                "h": CB.I_BSCDV_BEAM_H,
                "hf": CB.I_BSCDV_BEAM_T_HF,
                "bf": CB.I_BSCDV_BEAM_T_BF,
            },
            SH.I_BSST_CONCR_BEAM_I: {"b": CB.I_BSCDV_BEAM_B, "h": CB.I_BSCDV_BEAM_H},
        }

        LENGTH_SCALE = 1.0  # Robot API = metres
        if not bar.HasLabel(IRobotLabelType.I_LT_BAR_SECTION):
            raise ValueError("no section assigned")

        label = bar.GetLabel(IRobotLabelType.I_LT_BAR_SECTION)
        data = CastTo(label.Data, "IRobotBarSectionData")

        try:
            shape = SH(data.ShapeType)
            shape_name = shape.name
        except ValueError:
            shape = None
            shape_name = "UNKNOWN({})".format(data.ShapeType)
            raise ValueError(shape_name)


        result = {"name": label.Name, "shape": shape_name, "goe": {}}

        # rotation of the section (degrees), used only to know whether the
        # width seen in plan is b or h
        try:
            result["goe"]["gamma"] = data.GetValue(24)
        except Exception:
            result["goe"]["gamma"] = 0.0

        params = BEAM_PARAMS.get(shape)
        if params:
            concrete = CastTo(data.Concrete, "IRobotBarSectionConcreteData")
            for key, idx in params.items():
                result["goe"][key] = concrete.GetValue(int(idx)) * LENGTH_SCALE
        return result


# ======================================================================
# AUTOCAD: beams
# ======================================================================
class BeamsDrawer:
    """Draws the beams (from `Beams`) in AutoCAD, level by level, at the
    position given by PointLocator.pointxyz_to_cad_xy.

    Rule for a level L (elevation zL):
      * a beam is drawn in L when BOTH its end points are at zL
      * inclined beams (end points at different elevations) are not drawn
        in plan: they are listed in the console

    Each beam is drawn as a rectangle along its axis (from start node to end
    node, width = section width). Its name is written at the middle of the
    beam, along the beam direction (never upside down).
    """

    SECTION_LAYER = "Poutres"
    TEXT_LAYER = "Poutres - Textes"
    SECTION_COLOR = 5  # blue
    TEXT_COLOR = 4  # cyan

    # fields: {id}, {section} (Robot section name), {b}, {h}
    TEXT_FORMAT = "B{id} {section}"
    TEXT_HEIGHT = 0.15
    TEXT_FILL = 0.6  # text height <= TEXT_FILL x beam width

    DEFAULT_WIDTH = 0.20  # used when the section has no usable width (m)
    Z_TOL = 1e-3  # tolerance on elevations (m)
    MIN_LENGTH = 1e-3  # beams shorter than this are ignored (m)

    def __init__(self, beams, locator):
        self.beams = list(getattr(beams, "beams", beams))
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

        self.levels = self._levels()  # {level_name: elevation}
        self.drawn = []
        self.skipped = []  # (beam id, level, reason)

    # ---- levels / helpers ------------------------------------------------
    def _levels(self):
        levels = {}
        for (_grid, level), e in self.locator.index.items():
            levels.setdefault(level, e["z"])
        return dict(sorted(levels.items(), key=lambda kv: kv[1]))

    def _horizontal(self, beam):
        return (
            abs(beam["start_xyz"]["z"] - beam["end_xyz"]["z"]) <= self.Z_TOL
        )

    def beams_of_level(self, elevation):
        """Horizontal beams whose end points are both at `elevation`."""
        return [
            b
            for b in self.beams
            if self._horizontal(b)
            and abs(b["start_xyz"]["z"] - elevation) <= self.Z_TOL
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

    def _plan_width(self, goe):
        """Width of the beam seen in plan."""
        b, h = goe.get("b", 0), goe.get("h", 0)
        gamma = float(goe.get("gamma", 0.0) or 0.0)
        rotated = abs((gamma % 180.0) - 90.0) < 1.0  # section turned by 90 deg
        w = h if rotated else b
        return w if w > 0 else self.DEFAULT_WIDTH

    # ---- one beam --------------------------------------------------------
    def _draw_beam(self, beam, level_name, elevation):
        s, e = beam["start_xyz"], beam["end_xyz"]
        p1 = self._cad_xy(s["x"], s["y"], level_name, elevation)
        p2 = self._cad_xy(e["x"], e["y"], level_name, elevation)
        if p1 is None or p2 is None:
            self.skipped.append((beam["id"], level_name, "no CAD position"))
            return

        dx, dy = p2[0] - p1[0], p2[1] - p1[1]
        length = math.hypot(dx, dy)
        if length < self.MIN_LENGTH:
            self.skipped.append((beam["id"], level_name, "zero length in plan"))
            return

        section = beam["section"]
        goe = section.get("goe", {})
        if not (goe.get("b", 0) > 0 or goe.get("h", 0) > 0):
            print(
                "  Beam {}: no dimensions for shape {} -> default width {} m".format(
                    beam["id"], section.get("shape"), self.DEFAULT_WIDTH
                )
            )
        width = self._plan_width(goe)

        # outline: rectangle along the beam axis
        ux, uy = dx / length, dy / length
        nx, ny = -uy * width / 2.0, ux * width / 2.0
        pts = [
            (p1[0] + nx, p1[1] + ny),
            (p2[0] + nx, p2[1] + ny),
            (p2[0] - nx, p2[1] - ny),
            (p1[0] - nx, p1[1] - ny),
        ]
        self.factory.add_polygon(pts, self.section_layer)

        # name at the middle of the beam, along its direction, readable
        angle = math.atan2(dy, dx)
        if angle > math.pi / 2.0 + 1e-9 or angle <= -math.pi / 2.0 + 1e-9:
            angle = (angle + math.pi) % (2.0 * math.pi)  # never upside down
        label = self.TEXT_FORMAT.format(
            id=beam["id"],
            section=section.get("name", ""),
            b=goe.get("b", 0),
            h=goe.get("h", 0),
        )
        height = min(self.TEXT_HEIGHT, self.TEXT_FILL * width)
        self.factory.add_text(
            (p1[0] + p2[0]) / 2.0,
            (p1[1] + p2[1]) / 2.0,
            self.z_cad,
            label,
            height,
            self.text_layer,
            rotation=angle,
        )

        self.drawn.append(
            {
                "id": beam["id"],
                "level": level_name,
                "length": length,
                "width": width,
                "section": section.get("name"),
            }
        )

    # ---- public API ------------------------------------------------------
    def draw(self):
        for level_name, elevation in self.levels.items():
            beams = self.beams_of_level(elevation)
            print(
                "\nLevel '{}' (Z={}): {} beam(s)".format(
                    level_name, elevation, len(beams)
                )
            )
            for beam in beams:
                self._draw_beam(beam, level_name, elevation)

        drawn_ids = {d["id"] for d in self.drawn}
        inclined = [b["id"] for b in self.beams if not self._horizontal(b)]
        other = [
            b["id"]
            for b in self.beams
            if self._horizontal(b) and b["id"] not in drawn_ids
        ]
        if inclined:
            print("\nInclined beams (not drawn in plan): {}".format(inclined))
        if other:
            print("\nBeams not drawn (no level at their elevation): {}".format(other))
        for bid, level, reason in self.skipped:
            print("Skipped beam {} on {}: {}".format(bid, level, reason))

        self.session.refresh_view()
        return self


def BeamsPipeline(robot, locator):
    beams = Beams(robot)
    return BeamsDrawer(beams, locator).draw()


if __name__ == "__main__":
    robot = win32com.client.Dispatch("Robot.Application")
    robot.Visible = True

    # 1) grids: read from Robot and drawn in a new AutoCAD document
    locator = GridPipeline(robot)

    # 2) beams on the grids
    drawer = BeamsPipeline(robot, locator)

    print(
        "\nDone: {} beam(s) drawn, {} skipped.".format(
            len(drawer.drawn), len(drawer.skipped)
        )
    )

