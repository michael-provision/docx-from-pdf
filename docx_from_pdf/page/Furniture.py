'''Detect page headers and footers, the furniture repeated at the edges of pages.

A row of lines at the top (bottom) edge of a page belongs to the header (footer)
when only other furniture lies between it and that edge, and either

* the same row appears at the same place on another page (text compared
  exactly except that page numbers may differ, images by content), which also
  covers a bundle of documents that each carry their own letterhead, or
* it is set apart from the body by a clear gap and its group carries a page
  number such as ``Page 2 of 5``, which marks furniture on a single page.
'''

from copy import deepcopy
import hashlib
import re

from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from ..common.Collection import Collection
from ..image.ImageSpan import ImageSpan
from ..layout.Column import Column


PAGE_NUMBER_PATTERN = re.compile(
    r'^\s*(?:page\s*)?(\d+)\s*(?:(?:of|/)\s*(\d+))?\s*$'
    r'|^\s*page\s+(\d+)\b'
    r'|^\s*-\s*(\d+)\s*-\s*$',
    re.IGNORECASE)
# Two lines at the same place on different pages differ by less than this (pt).
POSITION_TOLERANCE = 3.0
# A page-number group needs at least this much space between it and the body,
# measured in heights of its own tallest row.
SEPARATION_IN_ROW_HEIGHTS = 2.0
# Furniture stays within this fraction of the page height from its edge.
MAX_BAND_RATIO = 0.25
# See ``footer_settle``.
FOOTER_SETTLE_RATIO = 0.01


class PageFurniture:
    '''Header and footer lines and shapes found on one page.'''

    def __init__(self):
        self.header_lines = Collection()
        self.footer_lines = Collection()
        self.header_shapes = Collection()
        self.footer_shapes = Collection()
        self.page_number_fields = False


def detect_furniture(raw_pages:list):
    '''Find the header and footer of every page.

    Args:
        raw_pages (list): Cleaned-up ``RawPage`` instances in page order.

    Returns:
        list: One ``PageFurniture`` per page.
    '''
    edges = [_edge_rows(raw_page) for raw_page in raw_pages]
    # compared line by line: rows on different pages can group the same lines differently
    page_lines = [[_line_key(line) for row in top + bottom for line in row] for top, bottom in edges]

    def is_repeated(row):
        if len(raw_pages) < 2: return False
        keys = [_line_key(line) for line in row]
        return sum(
            all(any(_same_line(key, other) for other in lines) for key in keys)
            for lines in page_lines
        ) >= 2

    result = []
    for raw_page, (top, bottom) in zip(raw_pages, edges):
        furniture = PageFurniture()
        for row in _furniture_rows(top, is_repeated, raw_page, from_top=True):
            furniture.header_lines.extend(row)
        for row in _furniture_rows(bottom, is_repeated, raw_page, from_top=False):
            furniture.footer_lines.extend(row)
        furniture.header_shapes = _band_shapes(raw_page, furniture.header_lines)
        furniture.footer_shapes = _band_shapes(raw_page, furniture.footer_lines)
        result.append(furniture)

    # page numbers become fields only where they agree with the document's own numbering
    numbered = [(index, _page_numbers(furniture)) for index, furniture in enumerate(result)]
    numbered = [(index, numbers) for index, numbers in numbered if numbers]
    consistent = bool(numbered) and all(
        current == index+1 and total in (None, len(raw_pages))
        for index, (current, total) in numbered
    )
    for index, _ in numbered:
        result[index].page_number_fields = consistent
    return result


def _edge_rows(raw_page):
    '''Rows of lines within the edge bands, ordered inwards from the top and the bottom edge.'''
    height = raw_page.height
    rows = Collection(raw_page.blocks).group_by_rows()
    top = [row for row in rows if row.bbox[3] <= height*MAX_BAND_RATIO]
    bottom = [row for row in rows if row.bbox[1] >= height*(1-MAX_BAND_RATIO)][::-1]
    return top, bottom


def _furniture_rows(rows:list, is_repeated, raw_page, from_top:bool):
    '''The outermost run of furniture rows, read inwards from the page edge.'''
    repeated = []
    for row in rows:
        if not is_repeated(row): break
        repeated.append(row)
    if repeated: return _without_body_overlap(repeated, raw_page)
    return _page_number_group(rows, raw_page, from_top)


def _page_number_group(rows:list, raw_page, from_top:bool):
    '''The edge group of rows if it carries a page number and stands clear of the body.'''
    group = []
    for row in rows:
        if group:
            previous = group[-1]
            gap = (row.bbox[1]-previous.bbox[3]) if from_top else (previous.bbox[1]-row.bbox[3])
            if gap > SEPARATION_IN_ROW_HEIGHTS*max(r.bbox.height for r in group): break
        group.append(row)
    if not any(_is_page_number(line.text) for row in group for line in row):
        return []
    grouped = {id(line) for row in group for line in row}
    body = [line for line in raw_page.blocks if id(line) not in grouped]
    if not body: return []
    if from_top:
        separation = min(line.bbox[1] for line in body) - group[-1].bbox[3]
    else:
        separation = group[-1].bbox[1] - max(line.bbox[3] for line in body)
    if separation < SEPARATION_IN_ROW_HEIGHTS*max(row.bbox.height for row in group):
        return []
    return _without_body_overlap(group, raw_page)


def _without_body_overlap(rows:list, raw_page):
    '''Drop inner rows until no body line shares vertical space with the furniture.'''
    selected = list(rows)
    while selected:
        lines = [line for row in selected for line in row]
        chosen = {id(line) for line in lines}
        y0 = min(line.bbox[1] for line in lines)
        y1 = max(line.bbox[3] for line in lines)
        if not any(
            line.bbox[1] < y1 and line.bbox[3] > y0
            for line in raw_page.blocks if id(line) not in chosen
        ):
            return selected
        selected.pop()
    return []


def _band_shapes(raw_page, lines):
    '''Shapes lying entirely within the band the furniture lines occupy.'''
    shapes = Collection()
    if not lines: return shapes
    y0 = min(line.bbox[1] for line in lines) - POSITION_TOLERANCE
    y1 = max(line.bbox[3] for line in lines) + POSITION_TOLERANCE
    for shape in raw_page.shapes:
        if shape.bbox[1] >= y0 and shape.bbox[3] <= y1:
            shapes.append(shape)
    return shapes


def _line_key(line):
    '''Position and content of a line: images by data, text exactly except page numbers.'''
    images = [span for span in line.spans if isinstance(span, ImageSpan)]
    if images:
        return line.bbox[0], line.bbox[1], hashlib.sha256(b''.join(span.image for span in images)).hexdigest()
    text = ' '.join(line.text.split()).lower()
    # only a page number may differ between pages; "Exhibit 1" and "Exhibit 2" are body titles
    if _is_page_number(text):
        text = re.sub(r'\d+', '#', text)
    return line.bbox[0], line.bbox[1], text


def _same_line(a:tuple, b:tuple):
    (x_a, y_a, content_a), (x_b, y_b, content_b) = a, b
    return content_a == content_b \
        and abs(x_a-x_b) <= POSITION_TOLERANCE and abs(y_a-y_b) <= POSITION_TOLERANCE


def _is_page_number(text:str):
    return bool(PAGE_NUMBER_PATTERN.match(text.strip()))


def _page_numbers(furniture:PageFurniture):
    '''The (current, total) page numbers the furniture shows, if any.'''
    for line in list(furniture.header_lines) + list(furniture.footer_lines):
        match = PAGE_NUMBER_PATTERN.match(line.text.strip())
        if not match: continue
        current = next(group for group in (match.group(1), match.group(3), match.group(4)) if group)
        total = match.group(2)
        return int(current), (int(total) if total else None)
    return None


class FurnitureContainer:
    '''Let block layouts write into a Word header or footer as into the document body.'''

    def __init__(self, part, width):
        self._part = part
        self._width = width

    def add_paragraph(self, *args, **kwargs):
        return self._part.add_paragraph(*args, **kwargs)

    def add_table(self, rows:int, cols:int):
        return self._part.add_table(rows, cols, self._width)

    @property
    def paragraphs(self):
        return self._part.paragraphs


def insert_page_number_fields(element):
    '''Replace page numbers in runs reading e.g. ``PAGE 1 OF 3`` with PAGE and NUMPAGES fields.

    Only a run whose whole text is a page number is changed, so the fields keep
    that run's formatting and pages added later still number themselves.
    '''
    for run in list(element.iter(qn('w:r'))):
        texts = run.findall(qn('w:t'))
        if len(texts) != 1 or run.getparent().tag == qn('w:fldSimple'): continue
        text = texts[0].text or ''
        match = PAGE_NUMBER_PATTERN.match(text)
        if not match: continue
        current = next(index for index in (1, 3, 4) if match.group(index))
        spans = [(match.span(current), 'PAGE')]
        if match.group(2): spans.append((match.span(2), 'NUMPAGES'))
        pieces, position = [], 0
        for (start, end), instruction in spans:
            pieces.append((text[position:start], None))
            pieces.append((text[start:end], instruction))
            position = end
        pieces.append((text[position:], None))
        for value, instruction in pieces:
            if not value: continue
            piece = deepcopy(run)
            for child in piece.findall(qn('w:t')): piece.remove(child)
            piece.append(_text_element(value))
            if instruction:
                field = OxmlElement('w:fldSimple')
                field.set(qn('w:instr'), f' {instruction} ')
                field.append(piece)
                piece = field
            run.addprevious(piece)
        run.getparent().remove(run)


def _text_element(value:str):
    element = OxmlElement('w:t')
    element.set(qn('xml:space'), 'preserve')
    element.text = value
    return element


def footer_settle(page_height:float, footer_bottom:float):
    '''How far below its source position the footer is placed, in pt.

    Rendered text rarely matches the PDF's metrics exactly, so a body that ends
    right above the footer can drift into it and push the page's end onto a
    blank page. Lowering the footer by about one percent of the page, and at
    most half its distance from the page edge, absorbs that drift.
    '''
    return min(page_height*FOOTER_SETTLE_RATIO, max(page_height-footer_bottom, 0)/2)


def remove_trailing_spacers(break_paragraph):
    '''Remove empty spacer paragraphs right before a page's section break.

    Layout adds one after a table that ends a section or page; at a page end
    the section-break paragraph already separates the table from what follows,
    and each spacer would need room the full page may not have.
    '''
    previous = break_paragraph.getprevious()
    while previous is not None and previous.tag == qn('w:p') \
        and previous.find(qn('w:r')) is None and previous.find(f"{qn('w:pPr')}/{qn('w:sectPr')}") is None:
        previous.getparent().remove(previous)
        previous = break_paragraph.getprevious()


class FurnitureLayout:
    '''A page header or footer: one layout per visual row, stacked top to bottom.

    Rows are laid out separately because furniture rows rarely flow into each
    other: a right-aligned page number above a left-aligned date would
    otherwise merge into one paragraph and lose both positions.
    '''

    def __init__(self, rows:list=None):
        self.rows = rows or []

    @classmethod
    def from_elements(cls, lines, shapes):
        '''Group furniture lines into rows, each with the shapes in its band.'''
        rows, top = [], None
        for row in Collection(lines).group_by_rows():
            x0, y0, x1, y1 = row.bbox
            column = Column().update_bbox((x0, y0 if top is None else top, x1, y1))
            column.add_elements(row)
            column.add_elements([
                shape for shape in shapes
                if y0 - POSITION_TOLERANCE <= (shape.bbox[1] + shape.bbox[3])/2 <= y1 + POSITION_TOLERANCE
            ])
            rows.append(column)
            top = y1
        return cls(rows)

    @property
    def bbox(self):
        return (
            min(row.bbox[0] for row in self.rows), self.rows[0].bbox[1],
            max(row.bbox[2] for row in self.rows), self.rows[-1].bbox[3],
        )

    def set_horizontal_bounds(self, x0:float, x1:float):
        '''Align every row to the page margins, as Word aligns header paragraphs.'''
        for row in self.rows:
            _, y0, _, y1 = row.bbox
            row.update_bbox((x0, y0, x1, y1))

    def parse(self, **settings):
        for row in self.rows: row.parse(**settings)

    def make_docx(self, container):
        for row in self.rows: row.make_docx(container)

    def store(self):
        return {'rows': [row.store() for row in self.rows]}

    @classmethod
    def restore(cls, data:dict):
        return cls([Column().restore(row) for row in data.get('rows', [])])
