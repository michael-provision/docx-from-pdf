"""PDFium-backed document and page primitives."""

from __future__ import annotations

from ctypes import byref, c_double, c_float, c_int, c_uint, create_string_buffer
from io import BytesIO
from math import cos, pi, sin, sqrt
from pathlib import Path

import pypdfium2 as pdfium
import pypdfium2.raw as pdfium_c

from ..common.geometry import Matrix, Rect


class PdfiumDocument:
    """Document wrapper exposing the subset of behavior used by the converter."""

    def __init__(self, pdf_file: str | None = None, password: str | None = None, stream: bytes | None = None):
        self._document = pdfium.PdfDocument(stream if stream is not None else Path(pdf_file), password=password or None)
        self.needs_pass = False

    def __len__(self):
        return len(self._document)

    def __iter__(self):
        for page_index in range(len(self)):
            yield self[page_index]

    def __getitem__(self, page_index: int):
        return PdfiumPage(self, page_index)

    def authenticate(self, password: str):
        return True

    def close(self):
        self._document.close()


class PdfiumPage:
    """Page wrapper with extraction helpers."""

    def __init__(self, document: PdfiumDocument, page_index: int):
        self.parent = document
        self.page_index = page_index
        self._page = document._document[page_index]
        self._pdf_cropbox = self._page.get_cropbox()
        crop_left, crop_bottom, crop_right, crop_top = self._pdf_cropbox
        self.width = crop_right - crop_left
        self.height = crop_top - crop_bottom
        self.rect = Rect(0.0, 0.0, self.width, self.height)
        self.cropbox = self.rect
        self.rotation = self._page.get_rotation()
        self.rotation_matrix = Matrix()

    def close(self):
        self._page.close()

    def extract_text_blocks(self, sort=None):
        text_page = self._page.get_textpage()
        chars = self._extract_chars(text_page)
        if sort:
            chars.sort(key=lambda char: (char["bbox"][1], char["bbox"][0]))
        lines = self._group_chars_into_lines(chars)
        return [{"type": 0, "bbox": self._bbox_for_lines(lines), "lines": lines}] if lines else []

    def extract_images(self, clip_image_res_ratio: float = 3.0):
        images = []
        for image_object in self._page.get_objects(filter=[pdfium_c.FPDF_PAGEOBJ_IMAGE], max_depth=3):
            bbox = self._pdf_bounds_to_page_rect(image_object.get_bounds())
            if bbox.get_area() <= 4:
                continue
            image_bytes, width, height = self.render_clip_to_png(
                bbox,
                zoom=clip_image_res_ratio,
                rm_text=True,
                rm_image=False,
            )
            images.append(
                {
                    "type": 1,
                    "bbox": tuple(bbox),
                    "width": width,
                    "height": height,
                    "image": image_bytes,
                }
            )
        return images

    def extract_paths(self):
        paths = []
        for path_object in self._page.get_objects(filter=[pdfium_c.FPDF_PAGEOBJ_PATH], max_depth=3):
            items = self._extract_path_items(path_object)
            if not items:
                continue
            fill_mode = c_int()
            stroke = c_int()
            pdfium_c.FPDFPath_GetDrawMode(path_object, byref(fill_mode), byref(stroke))
            path_type = ""
            if fill_mode.value:
                path_type += "f"
            if stroke.value:
                path_type += "s"
            if not path_type:
                continue
            stroke_width = c_float()
            pdfium_c.FPDFPageObj_GetStrokeWidth(path_object, byref(stroke_width))
            paths.append(
                {
                    "type": path_type,
                    "items": items,
                    "width": max(0.1, float(stroke_width.value)),
                    "color": self._get_object_color(path_object, stroke=True),
                    "fill": self._get_object_color(path_object, stroke=False),
                    "closePath": any(item[0] == "close" for item in items),
                }
            )
        return paths

    def render_clip_to_png(self, bbox=None, zoom: float = 3.0, rm_text: bool = False, rm_image: bool = False):
        clip = Rect(bbox) if bbox is not None else self.rect
        crop = (
            max(0.0, clip.x0),
            max(0.0, self.height - clip.y1),
            max(0.0, self.width - clip.x1),
            max(0.0, clip.y0),
        )
        deactivated_objects = self._deactivate_page_objects(rm_text=rm_text, rm_image=rm_image)
        try:
            bitmap = self._page.render(scale=zoom, crop=crop, rev_byteorder=True)
            pil_image = bitmap.to_pil()
            output = BytesIO()
            pil_image.save(output, format="PNG")
            return output.getvalue(), pil_image.width, pil_image.height
        finally:
            for page_object in deactivated_objects:
                pdfium_c.FPDFPageObj_SetIsActive(page_object, 1)

    def _extract_chars(self, text_page):
        chars = []
        for char_index in range(text_page.count_chars()):
            character = self._xml_compatible_text(text_page.get_text_range(char_index, 1))
            if not character:
                continue
            bbox = self._pdf_bounds_to_page_rect(text_page.get_charbox(char_index, loose=True))
            if bbox.is_empty:
                continue
            origin_x = c_double()
            origin_y = c_double()
            pdfium_c.FPDFText_GetCharOrigin(text_page, char_index, byref(origin_x), byref(origin_y))
            font_name = self._font_name(text_page, char_index)
            font_size = self._effective_font_size(text_page, char_index)
            chars.append(
                {
                    "c": character,
                    "bbox": tuple(bbox),
                    "origin": (float(origin_x.value), self.height - float(origin_y.value)),
                    "dir": self._text_direction(text_page, char_index),
                    "font": font_name,
                    "size": font_size,
                    "color": self._text_color(text_page, char_index),
                    "flags": self._font_flags(text_page, char_index, font_name),
                }
            )
        return chars

    def _group_chars_into_lines(self, chars):
        lines = []
        for char in chars:
            for line in reversed(lines[-8:]):
                if line["_dir"] != char["dir"]:
                    continue
                if abs(line["_center"] - self._line_center(char)) <= self._line_grouping_threshold(char):
                    line["_chars"].append(char)
                    line["_center"] = self._weighted_line_center(line["_chars"])
                    break
            else:
                lines.append({"_center": self._line_center(char), "_dir": char["dir"], "_chars": [char]})

        raw_lines = []
        for line in sorted(lines, key=self._line_sort_key):
            line_chars = self._sort_line_chars(line["_chars"], line["_dir"])
            for segment_chars in self._split_line_chars_by_gap(line_chars, line["_dir"]):
                self._add_missing_spaces(segment_chars, line["_dir"])
                spans = self._group_chars_into_spans(segment_chars)
                raw_lines.append({"wmode": 0, "dir": line["_dir"], "bbox": self._bbox_for_spans(spans), "spans": spans})
        return raw_lines

    @staticmethod
    def _split_line_chars_by_gap(chars, direction):
        if not chars:
            return []
        groups = [[chars[0]]]
        for char in chars[1:]:
            previous_char = groups[-1][-1]
            if direction == (0.0, -1.0):
                gap = previous_char["bbox"][1] - char["bbox"][3]
            else:
                gap = char["bbox"][0] - previous_char["bbox"][2]
            threshold = max(8.0, min(PdfiumPage._char_extent(previous_char), PdfiumPage._char_extent(char)) * 1.2)
            if gap > threshold:
                groups.append([char])
            else:
                groups[-1].append(char)
        return groups

    @staticmethod
    def _add_missing_spaces(chars, direction):
        for previous_char, char in zip(chars, chars[1:]):
            if previous_char["c"].isspace() or char["c"].isspace():
                continue
            if direction == (0.0, -1.0):
                gap = previous_char["bbox"][1] - char["bbox"][3]
            else:
                gap = char["bbox"][0] - previous_char["bbox"][2]
            threshold = max(1.5, min(PdfiumPage._char_advance_extent(previous_char), PdfiumPage._char_advance_extent(char)) * 0.35)
            if gap > threshold:
                previous_char["c"] += " "

    @staticmethod
    def _group_chars_into_spans(chars):
        spans = []
        current = None
        for char in chars:
            key = (char["font"], round(char["size"], 2), char["color"], char["flags"])
            if current is None or current["_key"] != key:
                current = {
                    "_key": key,
                    "font": char["font"],
                    "size": char["size"],
                    "color": char["color"],
                    "flags": char["flags"],
                    "chars": [],
                }
                spans.append(current)
            current["chars"].append({"c": char["c"], "bbox": char["bbox"], "origin": char["origin"]})
        for span in spans:
            span["bbox"] = PdfiumPage._bbox_for_chars(span["chars"])
            del span["_key"]
        return spans

    def _extract_path_items(self, path_object):
        matrix = path_object.get_matrix()
        items = []
        cursor = None
        bezier_points = []
        segment_count = pdfium_c.FPDFPath_CountSegments(path_object)
        for segment_index in range(segment_count):
            segment = pdfium_c.FPDFPath_GetPathSegment(path_object, segment_index)
            segment_type = pdfium_c.FPDFPathSegment_GetType(segment)
            x = c_float()
            y = c_float()
            pdfium_c.FPDFPathSegment_GetPoint(segment, byref(x), byref(y))
            point = self._matrix_point_to_page_point(matrix, x.value, y.value)
            if segment_type == pdfium_c.FPDF_SEGMENT_MOVETO:
                cursor = point
                bezier_points = []
            elif segment_type == pdfium_c.FPDF_SEGMENT_LINETO and cursor is not None:
                items.append(("l", cursor, point))
                cursor = point
                bezier_points = []
            elif segment_type == pdfium_c.FPDF_SEGMENT_BEZIERTO and cursor is not None:
                bezier_points.append(point)
                if len(bezier_points) == 3:
                    items.append(("c", cursor, bezier_points[0], bezier_points[1], bezier_points[2]))
                    cursor = bezier_points[2]
                    bezier_points = []
            if pdfium_c.FPDFPathSegment_GetClose(segment):
                items.append(("close",))
        return items

    def _deactivate_page_objects(self, rm_text: bool, rm_image: bool):
        if not rm_text and not rm_image:
            return []
        filters = []
        if rm_text:
            filters.append(pdfium_c.FPDF_PAGEOBJ_TEXT)
        if rm_image:
            filters.append(pdfium_c.FPDF_PAGEOBJ_IMAGE)
        deactivated_objects = []
        for page_object in self._page.get_objects(filter=filters, max_depth=3):
            active = c_int()
            pdfium_c.FPDFPageObj_GetIsActive(page_object, byref(active))
            if active.value:
                pdfium_c.FPDFPageObj_SetIsActive(page_object, 0)
                deactivated_objects.append(page_object)
        return deactivated_objects

    def _matrix_point_to_page_point(self, matrix, x: float, y: float):
        page_x = x * matrix.a + y * matrix.c + matrix.e
        page_y = x * matrix.b + y * matrix.d + matrix.f
        crop_left, _, _, crop_top = self._pdf_cropbox
        return (float(page_x - crop_left), float(crop_top - page_y))

    def _pdf_bounds_to_page_rect(self, bounds):
        left, bottom, right, top = bounds
        crop_left, _, _, crop_top = self._pdf_cropbox
        return Rect(left - crop_left, crop_top - top, right - crop_left, crop_top - bottom)

    @staticmethod
    def _effective_font_size(text_page, char_index: int):
        font_size = float(pdfium_c.FPDFText_GetFontSize(text_page, char_index) or 12.0)
        matrix = pdfium_c.FS_MATRIX()
        if not pdfium_c.FPDFText_GetMatrix(text_page, char_index, byref(matrix)):
            return font_size
        vertical_scale = sqrt(matrix.c * matrix.c + matrix.d * matrix.d)
        return font_size * vertical_scale if vertical_scale else font_size

    @staticmethod
    def _font_name(text_page, char_index: int):
        buffer = create_string_buffer(512)
        flags = c_int()
        length = pdfium_c.FPDFText_GetFontInfo(text_page, char_index, buffer, len(buffer), byref(flags))
        return buffer.value[:length].decode("utf-8", errors="ignore").split("+")[-1] or "Arial"

    @staticmethod
    def _font_flags(text_page, char_index: int, font_name: str):
        flags = 0
        if "Italic" in font_name or "Oblique" in font_name:
            flags |= 2**1
        if "Bold" in font_name or pdfium_c.FPDFText_GetFontWeight(text_page, char_index) >= 600:
            flags |= 2**4
        return flags

    @staticmethod
    def _text_direction(text_page, char_index: int):
        angle = float(pdfium_c.FPDFText_GetCharAngle(text_page, char_index)) % (2 * pi)
        if angle <= 0.1 or abs(angle - 2 * pi) <= 0.1:
            return (1.0, 0.0)
        if abs(angle - 1.5 * pi) <= 0.1:
            return (0.0, -1.0)
        return (round(cos(angle), 3), round(sin(angle), 3))

    @staticmethod
    def _text_color(text_page, char_index: int):
        red = c_uint()
        green = c_uint()
        blue = c_uint()
        alpha = c_uint()
        if not pdfium_c.FPDFText_GetFillColor(text_page, char_index, byref(red), byref(green), byref(blue), byref(alpha)):
            return 0
        return (red.value << 16) + (green.value << 8) + blue.value

    @staticmethod
    def _get_object_color(path_object, stroke: bool):
        red = c_uint()
        green = c_uint()
        blue = c_uint()
        alpha = c_uint()
        color_function = pdfium_c.FPDFPageObj_GetStrokeColor if stroke else pdfium_c.FPDFPageObj_GetFillColor
        if not color_function(path_object, byref(red), byref(green), byref(blue), byref(alpha)):
            return None
        return (red.value / 255.0, green.value / 255.0, blue.value / 255.0)

    @staticmethod
    def _xml_compatible_text(text: str):
        return "".join(char for char in text if PdfiumPage._is_xml_compatible_character(char))

    @staticmethod
    def _is_xml_compatible_character(char: str):
        codepoint = ord(char)
        return (
            codepoint in (0x9, 0xA, 0xD)
            or 0x20 <= codepoint <= 0xD7FF
            or 0xE000 <= codepoint <= 0xFFFD
            or 0x10000 <= codepoint <= 0x10FFFF
        )

    @staticmethod
    def _vertical_center(bbox):
        return (bbox[1] + bbox[3]) / 2.0

    @staticmethod
    def _horizontal_center(bbox):
        return (bbox[0] + bbox[2]) / 2.0

    @staticmethod
    def _line_center(char):
        if char["dir"] == (0.0, -1.0):
            return PdfiumPage._horizontal_center(char["bbox"])
        return PdfiumPage._vertical_center(char["bbox"])

    @staticmethod
    def _char_extent(char):
        if char["dir"] == (0.0, -1.0):
            return max(1.0, char["bbox"][2] - char["bbox"][0])
        return max(1.0, char["bbox"][3] - char["bbox"][1])

    @staticmethod
    def _char_advance_extent(char):
        if char["dir"] == (0.0, -1.0):
            return max(1.0, char["bbox"][3] - char["bbox"][1])
        return max(1.0, char["bbox"][2] - char["bbox"][0])

    @staticmethod
    def _line_grouping_threshold(char):
        return max(1.5, PdfiumPage._char_extent(char) * 0.35)

    @staticmethod
    def _weighted_line_center(chars):
        if chars[0]["dir"] == (0.0, -1.0):
            return sum(PdfiumPage._horizontal_center(char["bbox"]) * PdfiumPage._char_extent(char) for char in chars) / sum(
                PdfiumPage._char_extent(char) for char in chars
            )
        return sum(PdfiumPage._vertical_center(char["bbox"]) * PdfiumPage._char_extent(char) for char in chars) / sum(
            PdfiumPage._char_extent(char) for char in chars
        )

    @staticmethod
    def _line_sort_key(line):
        if line["_dir"] == (0.0, -1.0):
            return (min(char["bbox"][0] for char in line["_chars"]), min(char["bbox"][1] for char in line["_chars"]))
        return (line["_center"], min(char["bbox"][0] for char in line["_chars"]))

    @staticmethod
    def _sort_line_chars(chars, direction):
        if direction == (0.0, -1.0):
            return sorted(chars, key=lambda item: (-item["bbox"][1], item["bbox"][0]))
        return sorted(chars, key=lambda item: item["bbox"][0])

    @staticmethod
    def _bbox_for_chars(chars):
        bbox = Rect()
        for char in chars:
            bbox |= Rect(char["bbox"])
        return tuple(bbox)

    @staticmethod
    def _bbox_for_spans(spans):
        bbox = Rect()
        for span in spans:
            bbox |= Rect(span["bbox"])
        return tuple(bbox)

    @staticmethod
    def _bbox_for_lines(lines):
        bbox = Rect()
        for line in lines:
            bbox |= Rect(line["bbox"])
        return tuple(bbox)
