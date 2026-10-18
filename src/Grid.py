import json

import win32com.client
from win32com.client import CastTo


class RobotGrid:
    def __init__(self, robot) -> None:
        self.robot = robot
        self.grids = list()

    def LoadGrids(self):
        axe_manager = self.robot.Project.AxisMngr

        for grid_idx in range(1, axe_manager.Count + 1):
            print(grid_idx, axe_manager.Count)
            _grid = axe_manager.Get(int(grid_idx))
            if _grid is not None:
                grid = CastTo(
                    axe_manager.Get(grid_idx), "IRobotStructuralAxisGrid"
                )  ## why this return none dispte there is cylindrique grid
                if grid.Type == 1:
                    result = self.getGridCartesian(grid)
                    self.grids.append(result)

                if grid.Type == 2:
                    result = self.getGridPolar(grid)
                    self.grids[grid.Name] = result
                    self.grids["type"] = "Cylindrique"

    def getGridCartesian(self, gridobj):
        ## this works only for cartisien grid
        grid = CastTo(gridobj, "IRobotStructuralAxisGridCartesian")
        result = dict(name=grid.Name, X=dict(), Y=dict(), Z=dict())
        for i in range(
            1, grid.X.AxisCount + 1
        ):  # it show error that grid has no X prop
            axe = grid.X.GetAxis(i)
            result["X"][axe[1]] = axe[2]
        for i in range(1, grid.Y.AxisCount + 1):
            axe = grid.Y.GetAxis(i)
            result["Y"][axe[1]] = axe[2]
        for i in range(1, grid.Z.AxisCount + 1):
            axe = grid.Z.GetAxis(i)
            result["Z"][axe[1]] = axe[2]
        return result

    def getGridPolar(self, gridobj):
        # in the version 2015 api does not export the cylindiaue grid
        return dict()

    def __repr__(self) -> str:
        return json.dumps(self.grids, indent=4)


################################################
import json
import os
import re

import pythoncom
import win32com.client
from win32com.client import VARIANT


# ----------------------------------------------------------------------
# Data: one grid definition (name + X, Y axes + the levels it exists on)
# ----------------------------------------------------------------------
class GridData:
    """Grid definition: name + X, Y, Z axes ({label: coordinate})."""

    def __init__(self, name, x, y, z):
        self.name = name
        self.x = {str(k): float(v) for k, v in x.items()}
        self.y = {str(k): float(v) for k, v in y.items()}
        self.z = {str(k): float(v) for k, v in z.items()}
        self.validate()

    @classmethod
    def from_dict(cls, data):
        return cls(data["name"], data["X"], data["Y"], data["Z"])

    def validate(self):
        if not self.x or not self.y:
            raise ValueError(
                "Grid {!r}: X and Y must each contain at least one axis.".format(
                    self.name
                )
            )
        if not self.z:
            raise ValueError(
                "Grid {!r}: Z must contain at least one level.".format(self.name)
            )

    @property
    def xmin(self):
        return min(self.x.values())

    @property
    def xmax(self):
        return max(self.x.values())

    @property
    def ymin(self):
        return min(self.y.values())

    @property
    def ymax(self):
        return max(self.y.values())


# ----------------------------------------------------------------------
# Style: all drawing parameters in one place
# ----------------------------------------------------------------------
class GridStyle:
    def __init__(
        self,
        bubble_radius=0.20,
        text_height=0.20,
        title_height=0.50,
        title_gap=0.50,
        margin=1.0,
        bubble_gap=0.15,
        grid_spacing=3.0,
        level_spacing=4.0,
        window_padding=2.0,
        draw_points=True,
        points_visible=False,
        point_size=0.15,
        layer_color=8,
        window_color=7,
        point_color=1,
        line_type="CENTER",
        same_reference=True,
    ):
        self.bubble_radius = bubble_radius
        self.text_height = text_height
        self.title_height = title_height
        self.title_gap = title_gap
        self.margin = margin
        self.bubble_gap = bubble_gap
        self.grid_spacing = grid_spacing  # horizontal gap (same_reference=False only)
        self.level_spacing = level_spacing  # horizontal space between windows
        self.window_padding = window_padding  # free space around the content
        self.draw_points = draw_points  # create a POINT at every intersection
        self.points_visible = points_visible  # False: point layers are turned off
        self.point_size = point_size
        self.layer_color = layer_color
        self.window_color = window_color
        self.point_color = point_color
        self.line_type = line_type
        # True : all grids of a level share ONE translation (common origin)
        # False: grids of a level sit side by side, shifted along X
        self.same_reference = same_reference

    @property
    def offset(self):
        """Distance from the outer grid line to the bubble centre."""
        return self.margin + self.bubble_gap + self.bubble_radius

    @property
    def title_zone(self):
        """Height reserved at the top of each window for the level title."""
        return self.title_height + 2 * self.title_gap


# ----------------------------------------------------------------------
# AutoCAD connection
# ----------------------------------------------------------------------
class AutoCADSession:
    """Wraps the AutoCAD COM application and the *current* document."""

    def __init__(self, cad=None):
        if cad is None:
            self.app = win32com.client.GetActiveObject("AutoCAD.Application")
            self.app.Visible = True
        else:
            self.app = cad
        self._bind(self.app.ActiveDocument)

    def _bind(self, doc):
        self.doc = doc
        self.model = doc.ModelSpace
        self.layers = doc.Layers
        self.linetypes = doc.Linetypes

    def new_document(self):
        doc = self.app.Documents.Add()
        doc.Activate()
        self._bind(doc)
        return doc

    def set_title(self, title):
        try:
            self.doc.SummaryInfo.Title = title
        except Exception as exc:
            print("Could not set drawing title property: {}".format(exc))

    def save_document(self, path):
        self.doc.SaveAs(path)

    def ensure_linetype(self, name):
        try:
            self.linetypes.Item(name)
            return name
        except Exception:
            pass

        for filename in ("acad.lin", "acadiso.lin"):
            try:
                self.linetypes.Load(name, filename)
                return name
            except Exception:
                pass

        return "Continuous"

    def refresh_view(self):
        self.doc.SetVariable("LTSCALE", 1.0)
        self.doc.SetVariable("CELTSCALE", 1.0)
        self.doc.Regen(1)
        self.app.ZoomExtents()


# ----------------------------------------------------------------------
# Layers
# ----------------------------------------------------------------------
class LayerManager:
    def __init__(self, session, color):
        self.session = session
        self.color = color

    @staticmethod
    def safe_name(name):
        name = re.sub(r'[<>/\\":;?*|,=]', "-", str(name))
        name = name.strip().strip(".")
        return name[:255] or "Grid"

    def get_or_create(self, name, color=None, visible=True):
        name = self.safe_name(name)
        layers = self.session.layers

        try:
            layer = layers.Item(name)
        except Exception:
            try:
                layer = layers.Add(name)
            except Exception as exc:
                raise RuntimeError(
                    "Failed to create AutoCAD layer {!r}: {}".format(name, exc)
                ) from exc

        layer.Color = self.color if color is None else color
        layer.LayerOn = bool(visible)  # False = layer turned off (hidden)
        return name


# ----------------------------------------------------------------------
# Low-level entity creation (lines, bubbles, text, rectangle, points)
# ----------------------------------------------------------------------
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
        """Closed continuous rectangle (the level 'window')."""
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

    def add_point(self, x, y, z, layer_name):
        """AutoCAD POINT entity: snap to it with the NODE object snap."""
        pt = self.model.AddPoint(self.point(x, y, z))
        pt.Layer = layer_name
        pt.Color = self.BY_LAYER
        return pt


# ----------------------------------------------------------------------
# High-level drawer: ONE document, one window per level (left to right)
# ----------------------------------------------------------------------
class GridDrawer:
    COMMON_LEVEL = 0.0  # everything is drawn at this Z

    def __init__(
        self, grids, style=None, session=None, save_dir=None, doc_name="Grids"
    ):
        if isinstance(grids, GridData):
            grids = [grids]
        self.grids = list(grids)
        if not self.grids:
            raise ValueError("At least one grid is required.")

        names = [g.name for g in self.grids]
        if len(set(names)) != len(names):
            raise ValueError("Grid names must be unique: {}".format(names))

        self.style = style or GridStyle()
        self.session = session or AutoCADSession()
        self.save_dir = save_dir or os.path.join(
            os.path.expanduser("~"), "Documents", "grids"
        )
        self.doc_name = doc_name

        self.layers = LayerManager(self.session, self.style.layer_color)
        self.created_layers = []
        self.levels = self._group_by_level()

        # {(grid, level, x_label, y_label): (X, Y, Z)} in CAD world coordinates
        self.points = {}
        self.windows = {}  # {level: (left, bottom, right, top)}

    # -- grouping / sizes -----------------------------------------------
    def _group_by_level(self):
        levels = {}
        for grid in self.grids:
            for z_name in grid.z:
                levels.setdefault(z_name, []).append(grid)
        return levels

    def _extents(self, grid):
        """Local extents of a grid (bubble centres): (left, right, bottom, top)."""
        off = self.style.offset
        return (grid.xmin - off, grid.xmax + off, grid.ymin - off, grid.ymax + off)

    def _content_size(self, members):
        """(width, height) of the grids of one level."""
        exts = [self._extents(g) for g in members]
        ymin = min(e[2] for e in exts)
        ymax = max(e[3] for e in exts)
        height = ymax - ymin  # same in both modes: common baseline

        if self.style.same_reference:
            width = max(e[1] for e in exts) - min(e[0] for e in exts)
        else:
            width = sum(e[1] - e[0] for e in exts) + self.style.grid_spacing * (
                len(exts) - 1
            )
        return width, height

    def _window_size(self):
        """One common size, big enough for the largest level."""
        s = self.style
        sizes = [self._content_size(m) for m in self.levels.values()]
        width = max(w for w, _ in sizes) + 2 * s.window_padding
        height = max(h for _, h in sizes) + 2 * s.window_padding + s.title_zone
        return width, height

    # -- drawing (dx, dy = translation applied to the grid) -------------
    def _draw_x_axes(self, factory, grid, ext, dx, dy, z, layer):
        r = self.style.bubble_radius
        _, _, y_bottom, y_top = ext
        for label, x in grid.x.items():
            factory.add_line(
                (x + dx, y_bottom + r + dy, z), (x + dx, y_top - r + dy, z), layer
            )
            factory.add_bubble(x + dx, y_bottom + dy, z, label, layer)
            factory.add_bubble(x + dx, y_top + dy, z, label, layer)

    def _draw_y_axes(self, factory, grid, ext, dx, dy, z, layer):
        r = self.style.bubble_radius
        x_left, x_right, _, _ = ext
        for label, y in grid.y.items():
            factory.add_line(
                (x_left + r + dx, y + dy, z), (x_right - r + dx, y + dy, z), layer
            )
            factory.add_bubble(x_left + dx, y + dy, z, label, layer)
            factory.add_bubble(x_right + dx, y + dy, z, label, layer)

    def _draw_intersections(self, factory, grid, z_name, dx, dy, z):
        """Register (and optionally draw) every X/Y axis intersection.
        Points are created on layers that are turned off unless
        style.points_visible is True."""
        s = self.style
        elevation = grid.z[z_name]
        layer = None
        if s.draw_points:
            layer = self.layers.get_or_create(
                "Points - {} - {}".format(grid.name, z_name),
                s.point_color,
                visible=s.points_visible,
            )
            self.created_layers.append(layer)

        for x_label, x in grid.x.items():
            for y_label, y in grid.y.items():
                X, Y = x + dx, y + dy
                self.points[(grid.name, z_name, x_label, y_label)] = (X, Y, elevation)
                if layer:
                    factory.add_point(X, Y, z, layer)

    def _draw_window(
        self, factory, z_name, members, win_left, win_top, win_w, win_h, z
    ):
        s = self.style
        win_right = win_left + win_w
        win_bottom = win_top - win_h

        # Rectangle
        win_layer = self.layers.get_or_create("Windows", s.window_color)
        if win_layer not in self.created_layers:
            self.created_layers.append(win_layer)
        factory.add_rectangle(win_left, win_bottom, win_right, win_top, win_layer)
        self.windows[z_name] = (win_left, win_bottom, win_right, win_top)

        # Title, centred at the top of the window
        title_layer = self.layers.get_or_create("Titles - {}".format(z_name))
        self.created_layers.append(title_layer)
        cx = win_left + win_w / 2.0
        factory.add_text(
            cx, win_top - s.title_zone / 2.0, z, z_name, s.title_height, title_layer
        )

        # Content centred in the area below the title zone
        cy = win_top - s.title_zone - (win_h - s.title_zone) / 2.0
        content_w, _ = self._content_size(members)
        exts = {g.name: self._extents(g) for g in members}

        left_all = min(e[0] for e in exts.values())
        bottom_all = min(e[2] for e in exts.values())
        top_all = max(e[3] for e in exts.values())

        # Common vertical shift: same reference point for every grid
        dy = cy - (bottom_all + top_all) / 2.0

        if s.same_reference:
            # ONE translation for the whole level
            dx_common = (cx - content_w / 2.0) - left_all
        else:
            cursor = cx - content_w / 2.0

        for grid in members:
            ext = exts[grid.name]

            if s.same_reference:
                dx = dx_common
            else:
                dx = cursor - ext[0]  # only shifted along X
                cursor += (ext[1] - ext[0]) + s.grid_spacing

            layer = self.layers.get_or_create("{} - {}".format(grid.name, z_name))
            self.created_layers.append(layer)

            self._draw_x_axes(factory, grid, ext, dx, dy, z, layer)
            self._draw_y_axes(factory, grid, ext, dx, dy, z, layer)
            self._draw_intersections(factory, grid, z_name, dx, dy, z)
            print("  {!r}: shifted by ({:.2f}, {:.2f})".format(layer, dx, dy))

    # -- public API -----------------------------------------------------
    def draw(self):
        s = self.style
        self.session.new_document()
        line_type = self.session.ensure_linetype(s.line_type)
        factory = EntityFactory(self.session, s, line_type)

        # Make POINT entities visible (cross in a circle) when their layer is on
        self.session.doc.SetVariable("PDMODE", 35)
        self.session.doc.SetVariable("PDSIZE", s.point_size)

        win_w, win_h = self._window_size()
        print("Window size: {:.2f} x {:.2f}".format(win_w, win_h))

        z = self.COMMON_LEVEL
        win_left = 0.0
        for z_name, members in self.levels.items():
            print("\nLevel '{}': {}".format(z_name, ", ".join(g.name for g in members)))
            self._draw_window(factory, z_name, members, win_left, 0.0, win_w, win_h, z)
            win_left += win_w + s.level_spacing  # next level goes to the right

        self.session.set_title(self.doc_name)
        self.session.refresh_view()

        #         os.makedirs(self.save_dir, exist_ok=True)
        #         path = os.path.join(
        #             self.save_dir, LayerManager.safe_name(self.doc_name) + ".dwg"
        #         )
        #         self.session.save_document(path)
        #
        #         print("\nSaved: {}".format(path))
        #         print("Layers created:")
        #         for name in dict.fromkeys(self.created_layers):
        #             print(" - {}".format(name))
        return self

    def get_point(self, grid, level, x_label, y_label):
        """CAD coordinates (X, Y, Z) of the intersection of axis `x_label`
        (vertical line) and axis `y_label` (horizontal line) of `grid` on
        `level`. X, Y are the drawing coordinates, Z is the level elevation."""
        key = (grid, level, str(x_label), str(y_label))
        try:
            return self.points[key]
        except KeyError:
            raise KeyError(
                "No intersection {!r}: check grid, level and axis labels.".format(key)
            ) from None

    def export_points(self):
        rows = [
            {
                "grid": g,
                "level": lv,
                "x_axis": xl,
                "y_axis": yl,
                "X": p[0],
                "Y": p[1],
                "Z": p[2],
            }
            for (g, lv, xl, yl), p in self.points.items()
        ]
        return rows


# ----------------------------------------------------------------------
# Entry point
# ----------------------------------------------------------------------
class PointLocator:
    """Look up and reverse-search the intersection points registered by a
    GridDrawer (drawer.points = {(grid, level, x_label, y_label): (X, Y, Z)}).

    Forward : address (grid, level, x_label, y_label) -> CAD coordinates
    Reverse : CAD coordinates (X, Y, Z) -> axes passing through it, in every
              grid / level of the drawing
    """

    def __init__(self, drawer, tol=1e-6):
        self.drawer = drawer
        self.tol = tol
        self.index = self._build_index()

    # -- index ----------------------------------------------------------
    def _build_index(self):
        """{(grid, level): {"x": {label: X}, "y": {label: Y}, "z": elevation}}
        rebuilt from the registered points, so it always matches the drawing."""
        index = {}
        for (grid, level, x_label, y_label), (X, Y, Z) in self.drawer.points.items():
            entry = index.setdefault((grid, level), {"x": {}, "y": {}, "z": Z})
            entry["x"][x_label] = X  # vertical axis position
            entry["y"][y_label] = Y  # horizontal axis position
        return index

    # -- forward lookup -------------------------------------------------
    def get(self, grid, level, x_label, y_label):
        """Coordinates of one intersection, using drawer.get_point."""
        X, Y, Z = self.drawer.get_point(grid, level, x_label, y_label)
        return {
            "grid": grid,
            "level": level,
            "x_axis": str(x_label),
            "y_axis": str(y_label),
            "X": X,
            "Y": Y,
            "Z": Z,
        }

    def get_many(self, addresses):
        """addresses: iterable of (grid, level, x_label, y_label)."""
        return [self.get(*a) for a in addresses]

    # -- reverse lookup -------------------------------------------------
    def find_axes(self, x, y, z=None, tol=None):
        """Axes of ALL grids passing through the CAD point (x, y).

        z is the level elevation (optional). The geometry itself is drawn at
        Z = 0, so z only filters levels by their registered elevation.

        Returns one dict per (grid, level) that has at least one axis through
        the point:
          x_axes / y_axes  : labels of the vertical / horizontal axes through it
          is_intersection  : True when both an X axis and a Y axis pass through
        """
        tol = self.tol if tol is None else tol
        s = self.drawer.style
        off, r = s.offset, s.bubble_radius

        results = []
        for (grid, level), e in self.index.items():
            if z is not None and abs(e["z"] - z) > tol:
                continue

            xs, ys = e["x"], e["y"]
            # lines run between the bubbles, so test against the real extent
            x_lo, x_hi = min(xs.values()) - off + r, max(xs.values()) + off - r
            y_lo, y_hi = min(ys.values()) - off + r, max(ys.values()) + off - r

            x_axes = [
                label
                for label, X in xs.items()
                if abs(X - x) <= tol and y_lo - tol <= y <= y_hi + tol
            ]
            y_axes = [
                label
                for label, Y in ys.items()
                if abs(Y - y) <= tol and x_lo - tol <= x <= x_hi + tol
            ]
            if not x_axes and not y_axes:
                continue

            results.append(
                {
                    "grid": grid,
                    "level": level,
                    "elevation": e["z"],
                    "x_axes": x_axes,
                    "y_axes": y_axes,
                    "is_intersection": bool(x_axes and y_axes),
                }
            )
        return results

    def find_intersections(self, x, y, z=None, tol=None):
        """Same as find_axes, but only grids where an X and a Y axis cross
        exactly at the point."""
        return [r for r in self.find_axes(x, y, z, tol) if r["is_intersection"]]

    def axes_at_address(self, grid, level, x_label, y_label, tol=None):
        """Take a point by its address, then list the axes of every grid
        passing through that same CAD location."""
        X, Y, _ = self.drawer.get_point(grid, level, x_label, y_label)
        return self.find_axes(X, Y, None, tol)

    # -- display --------------------------------------------------------
    def describe(self, x, y, z=None, tol=None):
        found = self.find_axes(x, y, z, tol)
        if not found:
            return "No axis passes through ({}, {}).".format(x, y)
        lines = []
        for r in found:
            kind = "intersection" if r["is_intersection"] else "single axis"
            lines.append(
                "{} | {} (Z={}) | X axes: {} | Y axes: {} | {}".format(
                    r["grid"],
                    r["level"],
                    r["elevation"],
                    ", ".join(r["x_axes"]) or "-",
                    ", ".join(r["y_axes"]) or "-",
                    kind,
                )
            )
        return "\n".join(lines)

    def pointxyz_to_cad_xy(self, x, y, z, grid=None, only_inside=False, tol=None):
        """Structure coordinates (x, y, z) -> where to draw the point in AutoCAD.

        z must be the elevation of a level. Every level with that elevation is
        a candidate (each one is a separate window in the drawing). The
        translation (dx, dy) applied to a grid when it was drawn is recovered
        from the registered points, so nothing else has to change:
            dx = X_cad(axis) - x_axis_value
            dy = Y_cad(axis) - y_axis_value

        grid        : restrict the result to one grid name (default: all grids)
        only_inside : keep only grids whose extent contains (x, y)
        tol         : tolerance for matching z (default: self.tol)

        Returns a list of dicts, one per (grid, level) that matches:
            {"grid", "level", "elevation", "cad_x", "cad_y", "cad_z",
             "dx", "dy", "inside_grid", "window"}
        With same_reference=True all grids of a level share the same
        translation, so they all return the same cad_x / cad_y.
        """
        tol = self.tol if tol is None else tol
        x, y, z = float(x), float(y), float(z)
        grids = {g.name: g for g in self.drawer.grids}

        matches = [
            key
            for key, e in self.index.items()
            if abs(e["z"] - z) <= tol and (grid is None or key[0] == grid)
        ]
        if not matches:
            elevations = sorted({e["z"] for e in self.index.values()})
            raise ValueError(
                "No level at z={} (grid={!r}). Available elevations: {}".format(
                    z, grid, elevations
                )
            )

        results = []
        for g_name, level in matches:
            g = grids[g_name]
            e = self.index[(g_name, level)]

            # translation applied to this grid on this level
            x_label = next(iter(g.x))
            y_label = next(iter(g.y))
            dx = e["x"][x_label] - g.x[x_label]
            dy = e["y"][y_label] - g.y[y_label]

            inside = (
                g.xmin - tol <= x <= g.xmax + tol and g.ymin - tol <= y <= g.ymax + tol
            )
            if only_inside and not inside:
                continue

            results.append(
                {
                    "grid": g_name,
                    "level": level,
                    "elevation": e["z"],
                    "cad_x": x + dx,
                    "cad_y": y + dy,
                    "cad_z": self.drawer.COMMON_LEVEL,  # geometry is drawn at Z=0
                    "dx": dx,
                    "dy": dy,
                    "inside_grid": inside,
                    "window": self.drawer.windows.get(level),
                }
            )
        return results


def GridPipeline(robot):
    rgrid = RobotGrid(robot=robot)
    rgrid.LoadGrids()
    print(rgrid)
    grids = [GridData.from_dict(d) for d in rgrid.grids]

    # same_reference=True : grids of a level share one origin (overlaid)
    # same_reference=False: grids of a level sit side by side horizontally
    style = GridStyle(same_reference=True)
    drawer = GridDrawer(grids, style=style, doc_name="Grids").draw()
    return PointLocator(drawer)


if __name__ == "__main__":
    robot = win32com.client.Dispatch("Robot.Application")
    robot.Visible = True
    locator = GridPipeline(robot)
    # print(locator.get("Structure", "Level 2", "B", "A"))
    # print(locator.describe(3.0, 0.0))
# create class that takes points from the drawer and search using  get_point and also find the axes that intersection in all grids of given point address x, y , z
