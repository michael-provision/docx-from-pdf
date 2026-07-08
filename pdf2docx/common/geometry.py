"""Geometry primitives used by the PDF layout engine."""

from __future__ import annotations

from collections.abc import Iterable
from math import cos, radians, sin


class Matrix:
    """Affine transform matrix with the same tuple shape used by PDF libraries."""

    def __init__(self, *args):
        if len(args) == 0:
            self.a, self.b, self.c, self.d, self.e, self.f = 1.0, 0.0, 0.0, 1.0, 0.0, 0.0
        elif len(args) == 1:
            angle = float(args[0])
            theta = radians(angle)
            self.a = cos(theta)
            self.b = sin(theta)
            self.c = -sin(theta)
            self.d = cos(theta)
            self.e = 0.0
            self.f = 0.0
        elif len(args) == 2:
            self.a, self.b, self.c, self.d, self.e, self.f = (
                float(args[0]),
                0.0,
                0.0,
                float(args[1]),
                0.0,
                0.0,
            )
        elif len(args) == 6:
            self.a, self.b, self.c, self.d, self.e, self.f = [float(value) for value in args]
        else:
            raise ValueError("Matrix expects 0, 1, 2, or 6 values.")

    def __iter__(self):
        return iter((self.a, self.b, self.c, self.d, self.e, self.f))

    def __bool__(self):
        return tuple(self) != (1.0, 0.0, 0.0, 1.0, 0.0, 0.0)


class Point:
    """2D point supporting affine transform multiplication."""

    def __init__(self, x=0.0, y=None):
        if y is None and isinstance(x, Point):
            self.x = x.x
            self.y = x.y
        elif y is None and isinstance(x, Iterable) and not isinstance(x, (str, bytes)):
            values = list(x)
            self.x = float(values[0])
            self.y = float(values[1])
        else:
            self.x = float(x)
            self.y = float(y)

    def __iter__(self):
        return iter((self.x, self.y))

    def __getitem__(self, index):
        return (self.x, self.y)[index]

    def __mul__(self, matrix: Matrix):
        x = self.x * matrix.a + self.y * matrix.c + matrix.e
        y = self.x * matrix.b + self.y * matrix.d + matrix.f
        return Point(x, y)

    def __eq__(self, other):
        other_point = Point(other)
        return self.x == other_point.x and self.y == other_point.y

    def __repr__(self):
        return f"Point({self.x}, {self.y})"


class Rect:
    """Axis-aligned rectangle with union/intersection helpers."""

    def __init__(self, *args):
        if len(args) == 0:
            self.x0 = self.y0 = self.x1 = self.y1 = 0.0
        elif len(args) == 1:
            value = args[0]
            if isinstance(value, Rect):
                self.x0, self.y0, self.x1, self.y1 = value.x0, value.y0, value.x1, value.y1
            else:
                values = list(value)
                self.x0, self.y0, self.x1, self.y1 = [float(item) for item in values]
        elif len(args) == 4:
            self.x0, self.y0, self.x1, self.y1 = [float(item) for item in args]
        else:
            raise ValueError("Rect expects 0, 1, or 4 values.")

    @property
    def width(self):
        return max(0.0, self.x1 - self.x0)

    @property
    def height(self):
        return max(0.0, self.y1 - self.y0)

    @property
    def is_empty(self):
        return self.x1 <= self.x0 or self.y1 <= self.y0

    @property
    def tl(self):
        return Point(self.x0, self.y0)

    @property
    def br(self):
        return Point(self.x1, self.y1)

    def get_area(self):
        return self.width * self.height

    def intersects(self, other):
        other_rect = Rect(other)
        return not (
            self.x1 < other_rect.x0
            or other_rect.x1 < self.x0
            or self.y1 < other_rect.y0
            or other_rect.y1 < self.y0
        )

    def contains(self, other):
        other_rect = Rect(other)
        return (
            self.x0 <= other_rect.x0
            and self.y0 <= other_rect.y0
            and self.x1 >= other_rect.x1
            and self.y1 >= other_rect.y1
        )

    def __contains__(self, other):
        return self.contains(other)

    def __or__(self, other):
        other_rect = Rect(other)
        if not self:
            return Rect(other_rect)
        if not other_rect:
            return Rect(self)
        return Rect(
            min(self.x0, other_rect.x0),
            min(self.y0, other_rect.y0),
            max(self.x1, other_rect.x1),
            max(self.y1, other_rect.y1),
        )

    def __ior__(self, other):
        union = self | other
        self.x0, self.y0, self.x1, self.y1 = union.x0, union.y0, union.x1, union.y1
        return self

    def __and__(self, other):
        other_rect = Rect(other)
        if not self.intersects(other_rect):
            return Rect()
        return Rect(
            max(self.x0, other_rect.x0),
            max(self.y0, other_rect.y0),
            min(self.x1, other_rect.x1),
            min(self.y1, other_rect.y1),
        )

    def intersect(self, other):
        return self & other

    def __add__(self, value):
        dx0, dy0, dx1, dy1 = value
        return Rect(self.x0 + dx0, self.y0 + dy0, self.x1 + dx1, self.y1 + dy1)

    def __iadd__(self, value):
        expanded = self + value
        self.x0, self.y0, self.x1, self.y1 = expanded.x0, expanded.y0, expanded.x1, expanded.y1
        return self

    def __mul__(self, matrix: Matrix):
        points = [
            Point(self.x0, self.y0) * matrix,
            Point(self.x1, self.y0) * matrix,
            Point(self.x1, self.y1) * matrix,
            Point(self.x0, self.y1) * matrix,
        ]
        return Rect(
            min(point.x for point in points),
            min(point.y for point in points),
            max(point.x for point in points),
            max(point.y for point in points),
        )

    def __iter__(self):
        return iter((self.x0, self.y0, self.x1, self.y1))

    def __getitem__(self, index):
        return (self.x0, self.y0, self.x1, self.y1)[index]

    def __bool__(self):
        return (self.x0, self.y0, self.x1, self.y1) != (0.0, 0.0, 0.0, 0.0)

    def __repr__(self):
        return f"Rect({self.x0}, {self.y0}, {self.x1}, {self.y1})"
