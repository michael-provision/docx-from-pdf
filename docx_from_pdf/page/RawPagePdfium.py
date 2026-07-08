# -*- coding: utf-8 -*-

"""PDFium-backed raw page extraction."""

from .RawPage import RawPage
from ..common.Element import Element
from ..common.geometry import Rect
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

        shapes, images = self._preprocess_shapes(**settings)
        raw_dict["shapes"] = shapes

        image_blocks = self._preprocess_images(**settings)
        image_blocks = self._remove_images_covered_by_rendered_regions(image_blocks, images)
        raw_dict["blocks"].extend(image_blocks)
        raw_dict["blocks"].extend(images)

        raw_dict["shapes"].extend(self._preprocess_hyperlinks())
        Element.set_rotation_matrix(self.page_engine.rotation_matrix)

        return raw_dict

    def _preprocess_text(self, **settings):
        if settings["ocr"] == 1:
            raise SystemExit("OCR feature is planned but not implemented yet.")
        include = "hidden" if settings["ocr"] == 2 else "visible"
        return self.page_engine.extract_text_blocks(sort=settings.get("sort"), include=include)

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
        return self.page_engine.extract_hyperlinks()

    @staticmethod
    def _remove_images_covered_by_rendered_regions(image_blocks, rendered_regions):
        if not rendered_regions:
            return image_blocks

        rendered_bboxes = [Rect(image["bbox"]) for image in rendered_regions]
        filtered_images = []
        for image in image_blocks:
            image_bbox = Rect(image["bbox"])
            if any(
                RawPagePdfium._intersection_ratio(image_bbox, rendered_bbox) >= 0.95
                for rendered_bbox in rendered_bboxes
            ):
                continue
            filtered_images.append(image)
        return filtered_images

    @staticmethod
    def _intersection_ratio(source_bbox, target_bbox):
        source_area = source_bbox.get_area()
        if not source_area:
            return 0.0
        return (source_bbox & target_bbox).get_area() / source_area
