# -*- coding: utf-8 -*-

"""PDFium-backed raw page extraction."""

from .RawPage import RawPage
from ..common.Element import Element
from ..common.share import debug_plot
from ..image.ImagesExtractor import ImagesExtractor
from ..shape.Paths import Paths


class RawPagePdfium(RawPage):
    """Extract source contents from a PDFium page."""

    def extract_raw_dict(self, **settings):
        raw_dict = {}
        if not self.page_engine:
            return raw_dict

        self.width = self.page_engine.width
        self.height = self.page_engine.height
        raw_dict.update({"width": self.width, "height": self.height})

        text_blocks = self._preprocess_text(**settings)
        raw_dict["blocks"] = text_blocks

        image_blocks = self._preprocess_images(**settings)
        raw_dict["blocks"].extend(image_blocks)

        shapes, images = self._preprocess_shapes(**settings)
        raw_dict["shapes"] = shapes
        raw_dict["blocks"].extend(images)

        raw_dict["shapes"].extend(self._preprocess_hyperlinks())
        Element.set_rotation_matrix(self.page_engine.rotation_matrix)

        return raw_dict

    def _preprocess_text(self, **settings):
        if settings["ocr"] == 1:
            raise SystemExit("OCR feature is planned but not implemented yet.")
        if settings["ocr"] == 2:
            return []
        return self.page_engine.extract_text_blocks(sort=settings.get("sort"))

    def _preprocess_images(self, **settings):
        if settings["ocr"] == 2:
            return []
        return ImagesExtractor(self.page_engine).extract_images(settings["clip_image_res_ratio"])

    def _preprocess_shapes(self, **settings):
        paths = self._init_paths(**settings)
        return paths.to_shapes_and_images(
            settings["min_svg_gap_dx"],
            settings["min_svg_gap_dy"],
            settings["min_svg_w"],
            settings["min_svg_h"],
            settings["clip_image_res_ratio"],
        )

    @debug_plot("Source Paths")
    def _init_paths(self, **settings):
        return Paths(parent=self).restore(self.page_engine.extract_paths())

    def _preprocess_hyperlinks(self):
        return []
