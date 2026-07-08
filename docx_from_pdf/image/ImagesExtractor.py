"""Extract images and rendered vector regions from PDF pages."""

from ..common.share import BlockType
from ..common.algorithm import recursive_xy_cut, inner_contours, xy_project_profile


class ImagesExtractor:
    """Extract images from a backend page."""

    def __init__(self, page) -> None:
        self._page = page

    def clip_page_to_pixmap(self, bbox=None, rm_image: bool = False, zoom: float = 3.0):
        image_bytes, width, height = self._page.render_clip_to_png(
            bbox=bbox,
            zoom=zoom,
            rm_text=True,
            rm_image=rm_image,
        )
        return RenderedImage(image_bytes=image_bytes, width=width, height=height)

    def clip_page_to_dict(
        self,
        bbox=None,
        rm_image: bool = False,
        clip_image_res_ratio: float = 3.0,
    ):
        pixmap = self.clip_page_to_pixmap(
            bbox=bbox, rm_image=rm_image, zoom=clip_image_res_ratio
        )
        return self._to_raw_dict(pixmap, bbox)

    def extract_images(self, clip_image_res_ratio: float = 3.0):
        return self._page.extract_images(clip_image_res_ratio)

    def detect_svg_contours(
        self, min_svg_gap_dx: float, min_svg_gap_dy: float, min_w: float, min_h: float
    ):
        import cv2 as cv

        pixmap = self.clip_page_to_pixmap(rm_image=True, zoom=1.0)
        src = self._pixmap_to_cv_image(pixmap)

        gray = cv.cvtColor(src, cv.COLOR_BGR2GRAY)
        _, binary = cv.threshold(gray, 253, 255, cv.THRESH_BINARY_INV)

        external_bboxes = recursive_xy_cut(
            binary, min_dx=min_svg_gap_dx, min_dy=min_svg_gap_dy
        )
        grouped_inner_bboxes = [
            inner_contours(binary, bbox, min_w, min_h) for bbox in external_bboxes
        ]
        groups = list(zip(external_bboxes, grouped_inner_bboxes))

        debug = False
        if debug:
            for i, (x0, y0, x1, y1) in enumerate(external_bboxes):
                arr = xy_project_profile(src[y0:y1, x0:x1, :], binary[y0:y1, x0:x1])
                cv.imshow(f"sub-image-{i}", arr)

            for bbox, inner_bboxes in groups:
                x0, y0, x1, y1 = bbox
                cv.rectangle(src, (x0, y0), (x1, y1), (255, 0, 0), 1)

                for u0, v0, u1, v1 in inner_bboxes:
                    cv.rectangle(src, (u0, v0), (u1, v1), (0, 0, 255), 1)

            cv.imshow("img", src)
            cv.waitKey(0)

        return groups

    @staticmethod
    def _to_raw_dict(image, bbox):
        return {
            "type": BlockType.IMAGE.value,
            "bbox": tuple(bbox),
            "width": image.width,
            "height": image.height,
            "image": image.tobytes(),
        }

    @staticmethod
    def _rotate_image(pixmap, rotation: int):
        import cv2 as cv
        import numpy as np

        img = ImagesExtractor._pixmap_to_cv_image(pixmap)
        h, w = img.shape[:2]
        x0, y0 = w // 2, h // 2
        matrix = cv.getRotationMatrix2D((x0, y0), rotation, 1.0)
        cos = np.abs(matrix[0, 0])
        sin = np.abs(matrix[0, 1])
        width = int((h * sin) + (w * cos))
        height = int((h * cos) + (w * sin))
        matrix[0, 2] += (width / 2) - x0
        matrix[1, 2] += (height / 2) - y0
        rotated_img = cv.warpAffine(img, matrix, (width, height))
        _, im_png = cv.imencode(".png", rotated_img)
        return im_png.tobytes()

    @staticmethod
    def _pixmap_to_cv_image(pixmap):
        import cv2 as cv
        import numpy as np

        return cv.imdecode(np.frombuffer(pixmap.tobytes(), np.uint8), cv.IMREAD_COLOR)


class RenderedImage:
    def __init__(self, image_bytes: bytes, width: int, height: int):
        self.image_bytes = image_bytes
        self.width = width
        self.height = height

    def tobytes(self):
        return self.image_bytes
