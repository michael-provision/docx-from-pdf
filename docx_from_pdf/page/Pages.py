# -*- coding: utf-8 -*-

'''Collection of :py:class:`~docx_from_pdf.page.Page` instances.'''

import logging

from .Furniture import FurnitureLayout, detect_furniture, footer_settle
from .RawPageFactory import RawPageFactory
from ..common.Collection import BaseCollection
from ..font.Fonts import Fonts


class Pages(BaseCollection):
    '''A collection of ``Page``.'''

    def parse(self, pdf_doc, **settings):
        '''Analyze document structure, e.g. page section, header, footer.

        Args:
            pdf_doc: PDF document instance.
            settings (dict): Parsing parameters.
        '''
        # ---------------------------------------------
        # 0. extract fonts properties, especially line height ratio
        # ---------------------------------------------
        fonts = Fonts.extract(pdf_doc)

        # ---------------------------------------------
        # 1. extract and then clean up raw page
        # ---------------------------------------------
        pages, raw_pages = [], []
        words_found = False
        for page in self:
            if page.skip_parsing: continue

            # init and extract data from PDF
            raw_page = RawPageFactory.create(page_engine=pdf_doc[page.id], backend="pdfium")
            raw_page.restore(**settings)

            # check if any words are extracted since scanned pdf may be directed
            if not words_found and raw_page.raw_text.strip():
                words_found = True

            # process blocks and shapes based on bbox
            raw_page.clean_up(**settings)

            # process font properties
            raw_page.process_font(fonts)            

            # after this step, we can get some basic properties
            # NOTE: floating images are detected when cleaning up blocks, so collect them here
            page.width = raw_page.width
            page.height = raw_page.height
            page.float_images.reset().extend(raw_page.blocks.floating_image_blocks)

            raw_pages.append(raw_page)
            pages.append(page)

        # show message if no words found
        if not words_found:
            logging.warning('Words count: 0. It might be a scanned pdf, which is not supported yet.')

        
        # ---------------------------------------------
        # 2. parse structure in document/pages level
        # ---------------------------------------------
        # NOTE: blocks structure might be changed in this step, e.g. promote page header/footer,
        # so blocks structure based process, e.g. calculating margin, parse section should be
        # run after this step.
        furniture = detect_furniture(raw_pages)


        # ---------------------------------------------
        # 3. parse structure in page level, e.g. page margin, section
        # ---------------------------------------------
        # parse sections
        for page, raw_page, page_furniture in zip(pages, raw_pages, furniture):
            # left/right margins still cover the header and footer, which align to them
            left, right, _, _ = raw_page.calculate_margin(**settings)
            Pages._promote_furniture(page, raw_page, page_furniture)
            _, _, top, bottom = raw_page.calculate_margin(**settings)
            if page.header:
                top = max(top, page.header.bbox[3])
            if page.footer:
                settle = footer_settle(raw_page.height, page.footer.bbox[3])
                bottom = max(bottom, raw_page.height - page.footer.bbox[1] - settle)
            raw_page.margin = page.margin = (left, right, top, bottom)
            for layout in (page.header, page.footer):
                if layout:
                    layout.set_horizontal_bounds(left, raw_page.width-right)

            # page section
            sections = raw_page.parse_section(**settings)
            page.sections.extend(sections)


    @staticmethod
    def _promote_furniture(page, raw_page, furniture):
        '''Move header and footer lines and shapes out of the page body into their own layouts.'''
        moved = {id(element) for element in (
            list(furniture.header_lines) + list(furniture.footer_lines)
            + list(furniture.header_shapes) + list(furniture.footer_shapes))}
        page.page_number_fields = furniture.page_number_fields
        for name, lines, shapes in (
            ('header', furniture.header_lines, furniture.header_shapes),
            ('footer', furniture.footer_lines, furniture.footer_shapes),
        ):
            if not lines:
                continue
            setattr(page, name, FurnitureLayout.from_elements(lines, shapes))
        raw_page.blocks.reset([line for line in raw_page.blocks if id(line) not in moved])
        raw_page.shapes.reset([shape for shape in raw_page.shapes if id(shape) not in moved])
