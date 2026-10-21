'''
The test framework: pytest, pytest-cov.

We have a set of PDF files as test inputs.

For a test file foo.pdf, we convert it into a file foo.pdf.docx.

To check whether this has worked as expected, we use Python package docx2pdf
(which uses Word) on Windows, or Libreoffice command line on other platforms,
to convert foo.pdf.docx into foo.pdf.docx.pdf.

We then compare foo.pdf.docx.pdf with the original foo.pdf file using opencv,
generating a similarity value.

So on Windows we require Word is installed, and on other platforms we require
that Libreoffice is installed.

If docx2pdf fails with `Object reference not set to an instance of an
object. Did not convert`, it might be necessary to follow the instructions at:

    https://stackoverflow.com/questions/24860351/object-reference-not-set-to-an-instance-of-an-object-did-not-convert

    In a Cmd window run:
        DCOMCNFG
    Then:
        Console Root > Component Services > Computers > My Computer > DCOM Config > Microsoft Word 97 - 2003 Document
    Then: Right click then properties then Identity tab and set a username and
    password.
'''

import glob
import os
import io
import base64
from io import BytesIO
import numpy as np
import cv2 as cv
import pypdfium2 as pdfium
from docx import Document
from docx_from_pdf import Converter, parse
import subprocess
import time
import shutil
import platform
import pytest

from docx.oxml.ns import qn
from docx.shared import Pt

from docx_from_pdf.backend.pdfium import PdfiumDocument
from docx_from_pdf.common import docx as docx_helpers
from docx_from_pdf.common.share import BlockType, rgb_value
from docx_from_pdf.page.RawPagePdfium import RawPagePdfium
from docx_from_pdf.shape.Path import Path
from docx_from_pdf.table.Cell import Cell
from docx_from_pdf.table.Row import Row
from docx_from_pdf.text.TextBlock import TextBlock


root_path = os.path.abspath(f'{__file__}/../..')
script_path = os.path.abspath(__file__) # current script path
test_dir = os.path.dirname(script_path)
sample_path = os.path.join(test_dir, 'samples')
output_path = os.path.join(test_dir, 'outputs')


def build_pages_pdf_bytes(content_streams):
    '''A letter-size PDF with one Helvetica page per content stream.'''
    count = len(content_streams)
    kids = ' '.join(f'{3 + 2*i} 0 R' for i in range(count))
    objects = [
        b'<< /Type /Catalog /Pages 2 0 R >>',
        f'<< /Type /Pages /Kids [{kids}] /Count {count} >>'.encode(),
    ]
    font = 3 + 2*count
    for index, stream in enumerate(content_streams):
        objects.append(
            f'<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 {font} 0 R >> >> /Contents {4 + 2*index} 0 R >>'.encode())
        objects.append(b'<< /Length ' + str(len(stream)).encode() + b' >>\nstream\n' + stream + b'\nendstream')
    objects.append(b'<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>')
    pdf = bytearray(b'%PDF-1.4\n')
    offsets = [0]
    for index, obj in enumerate(objects, start=1):
        offsets.append(len(pdf))
        pdf.extend(f'{index} 0 obj\n'.encode())
        pdf.extend(obj)
        pdf.extend(b'\nendobj\n')
    xref_offset = len(pdf)
    pdf.extend(f'xref\n0 {len(objects) + 1}\n'.encode())
    pdf.extend(b'0000000000 65535 f \n')
    for offset in offsets[1:]:
        pdf.extend(f'{offset:010d} 00000 n \n'.encode())
    pdf.extend(f'trailer\n<< /Size {len(objects) + 1} /Root 1 0 R >>\nstartxref\n{xref_offset}\n%%EOF\n'.encode())
    return bytes(pdf)


def text_at(x, y, text, size=10):
    return f'BT /F1 {size} Tf {x} {y} Td ({text}) Tj ET '.encode()


def convert_pages(tmp_name, content_streams):
    pdf_file = os.path.join(output_path, f'{tmp_name}.pdf')
    docx_file = os.path.join(output_path, f'{tmp_name}.docx')
    with open(pdf_file, 'wb') as f:
        f.write(build_pages_pdf_bytes(content_streams))
    cv = Converter(pdf_file)
    try:
        cv.convert(docx_file, raw_exceptions=True)
    finally:
        cv.close()
    return Document(docx_file)


def body_text(document):
    return ' '.join(paragraph.text for paragraph in document.paragraphs)


def part_text(part):
    return ' '.join(''.join(t.text or '' for t in paragraph._p.iter(qn('w:t'))) for paragraph in part.paragraphs)


def page_fields(part):
    return [field.get(qn('w:instr')).strip() for field in part._element.iter(qn('w:fldSimple'))]


def build_pdf_bytes(content_stream):
    objects = [
        b'<< /Type /Catalog /Pages 2 0 R >>',
        b'<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
        b'<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>',
        b'<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>',
        b'<< /Length ' + str(len(content_stream)).encode() + b' >>\nstream\n' + content_stream + b'\nendstream',
    ]
    pdf = bytearray(b'%PDF-1.4\n')
    offsets = [0]
    for index, obj in enumerate(objects, start=1):
        offsets.append(len(pdf))
        pdf.extend(f'{index} 0 obj\n'.encode())
        pdf.extend(obj)
        pdf.extend(b'\nendobj\n')

    xref_offset = len(pdf)
    pdf.extend(f'xref\n0 {len(objects) + 1}\n'.encode())
    pdf.extend(b'0000000000 65535 f \n')
    for offset in offsets[1:]:
        pdf.extend(f'{offset:010d} 00000 n \n'.encode())
    pdf.extend(b'trailer\n')
    pdf.extend(f'<< /Size {len(objects) + 1} /Root 1 0 R >>\n'.encode())
    pdf.extend(b'startxref\n')
    pdf.extend(f'{xref_offset}\n%%EOF\n'.encode())
    return bytes(pdf)


def assert_schema_order(parent, order):
    tags = [qn(name) for name in order]
    positions = [tags.index(child.tag) for child in parent if child.tag in tags]
    assert positions == sorted(positions), [child.tag for child in parent]


def build_table_cell_with_two_text_rows(border_width=(0, 0, 0, 0)):
    return Cell({
        'bbox': (0, 0, 400, 100),
        'border_width': border_width,
        'blocks': [{
            'type': BlockType.TEXT.value,
            'lines': [
                {
                    'bbox': (10, 10, 220, 30),
                    'wmode': 0,
                    'dir': (1.0, 0.0),
                    'spans': [{
                        'bbox': (10, 10, 220, 30),
                        'font': 'Arial',
                        'size': 12.0,
                        'color': 0,
                        'flags': 0,
                        'chars': [{'c': 'First line', 'bbox': (10, 10, 220, 30)}],
                    }],
                },
                {
                    'bbox': (10, 34, 200, 46),
                    'wmode': 0,
                    'dir': (1.0, 0.0),
                    'spans': [{
                        'bbox': (10, 34, 200, 46),
                        'font': 'Arial',
                        'size': 10.0,
                        'color': 0,
                        'flags': 0,
                        'chars': [{'c': 'Second line', 'bbox': (10, 34, 200, 46)}],
                    }],
                },
            ],
        }],
    })


def build_mixed_image_text_block():
    png_bytes = base64.b64decode(
        'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII='
    )
    return TextBlock({
        'line_space': 10.0,
        'line_space_type': 0,
        'tab_stops': [123.0],
        'lines': [
            {
                'bbox': (54, 43, 172, 68),
                'wmode': 0,
                'dir': (1.0, 0.0),
                'spans': [{
                    'bbox': (54, 43, 172, 68),
                    'width': 118,
                    'height': 25,
                    'image': png_bytes,
                }],
            },
            {
                'bbox': (177, 61, 490, 71),
                'wmode': 0,
                'dir': (1.0, 0.0),
                'spans': [{
                    'bbox': (177, 61, 490, 71),
                    'font': 'Arial',
                    'size': 10.0,
                    'color': 0,
                    'flags': 0,
                    'text': 'Subtitle',
                }],
            },
        ],
    })


def get_page_similarity(page_a, page_b, diff_img_filename='diff.png'):
    '''Calculate page similarity index: [0, 1].'''
    # page to opencv image
    image_a = get_page_image(page_a)
    image_b = get_page_image(page_b)

    # resize if different shape
    if image_a.shape != image_b.shape:
        rows, cols = image_a.shape[:2]
        image_b = cv.resize(image_b, (cols, rows))

    # write different image
    if diff_img_filename:
        diff = cv.subtract(image_a, image_b)
        cv.imwrite(diff_img_filename, diff)

    return get_mssism(image_a, image_b)


def get_page_image(pdf_page):
    '''Convert PDF page to opencv image.'''
    bitmap = pdf_page.render(rev_byteorder=True)
    image = bitmap.to_pil()
    output = BytesIO()
    image.save(output, format='PNG')
    img_byte = output.getvalue()
    img = np.frombuffer(img_byte, np.uint8)
    return cv.imdecode(img, cv.IMREAD_COLOR)


def get_mssism(i1, i2, kernel=(15,15)):
    '''Calculate mean Structural Similarity Index (SSIM).
    https://docs.opencv.org/4.x/d5/dc4/tutorial_video_input_psnr_ssim.html
    '''
    C1 = 6.5025
    C2 = 58.5225
    # INITS
    I1 = np.float32(i1) # cannot calculate on one byte large values
    I2 = np.float32(i2)
    I2_2 = I2 * I2 # I2^2
    I1_2 = I1 * I1 # I1^2
    I1_I2 = I1 * I2 # I1 * I2
    # END INITS
    # PRELIMINARY COMPUTING
    mu1 = cv.GaussianBlur(I1, kernel, 1.5)
    mu2 = cv.GaussianBlur(I2, kernel, 1.5)
    mu1_2 = mu1 * mu1
    mu2_2 = mu2 * mu2
    mu1_mu2 = mu1 * mu2
    sigma1_2 = cv.GaussianBlur(I1_2, kernel, 1.5)
    sigma1_2 -= mu1_2
    sigma2_2 = cv.GaussianBlur(I2_2, kernel, 1.5)
    sigma2_2 -= mu2_2
    sigma12 = cv.GaussianBlur(I1_I2, kernel, 1.5)
    sigma12 -= mu1_mu2
    t1 = 2 * mu1_mu2 + C1
    t2 = 2 * sigma12 + C2
    t3 = t1 * t2                    # t3 = ((2*mu1_mu2 + C1).*(2*sigma12 + C2))
    t1 = mu1_2 + mu2_2 + C1
    t2 = sigma1_2 + sigma2_2 + C2
    t1 = t1 * t2                    # t1 =((mu1_2 + mu2_2 + C1).*(sigma1_2 + sigma2_2 + C2))
    ssim_map = cv.divide(t3, t1)    # ssim_map =  t3./t1;
    mssim = cv.mean(ssim_map)       # mssim = average of ssim map
    return np.mean(mssim[0:3])


def run(command):
   print(f'Running: {command}')
   subprocess.run(command, shell=1, check=1)


def document_to(in_, out):
    if platform.system() == 'Windows':
        return word_to(in_, out)
    else:
        return libreoffice_to(in_, out)


_g_word_to_docx2pdf = False

def word_to(in_, out):
    global _g_word_to_docx2pdf
    if not _g_word_to_docx2pdf:
        run('pip install docx2pdf')
        import docx2pdf
        _g_word_to_docx2pdf = True
    assert os.path.isfile(in_), f'Not a file: {in_=}'
    run(f'docx2pdf {in_} {out}')
    return
    import docx2pdf
    try:
        docx2pdf.convert(in_, out)
    except Exception as e:
        print(f'docx2pdf.convert() raised exception: {e}')
        raise
    


def libreoffice_to(in_, out):
    '''Converts file to pdf using libreoffice. Returns generated path
    f'{in_}.pdf'.'''
    # Libreoffice does not allow direct specification of the output path and
    # goes wrong wtih paths with multiple '.' characters, so we work on a
    # temporary. Also it does not return non-zero if it fails so we check
    # mtime.
    #print(f'{in_=} {out=}')
    assert os.path.isfile(in_)
    in_root, in_ext = os.path.splitext(in_)
    _, out_ext = os.path.splitext(out)
    out_dir = os.path.dirname(out)
    temp = f'{out_dir}/_temp_libreoffice_to'
    in2 = f'{temp}{in_ext}'
    out2 = f'{temp}{out_ext}'
    shutil.copy2(in_, in2)
    try:
        t = time.time()
        #print(f'{in_=} {in2=} {in_ext=}')
        run(f'libreoffice --convert-to {out_ext[1:]} --outdir {out_dir} {in2}')
        os.rename(out2, out)
        t_out = os.path.getmtime(out)
        assert t_out >= t, f'libreoffice failed to update/create {out=}'
    finally:
        os.remove(in2)
        if os.path.isfile(out2):
            os.remove(out2)


def compare_pdf(pdf1, pdf2, num_pages=None):
    #print(f'Comparing {pdf1=} {pdf2=}')
    doc1 = pdfium.PdfDocument(pdf1)
    doc2 = pdfium.PdfDocument(pdf2)
    try:
        if num_pages:
            n1 = num_pages
        else:
            n1 = len(doc1)
            n2 = len(doc2)
            if n1 != n2:
                print(f'Differing numbers of pages: {n1=} {n2=}.')
                return -1
        sidx = 0
        # Find average similarity.
        for n in range(n1):
            diff_png = f'{pdf2}.diff.{n}.png'
            sidx_n = get_page_similarity(doc1[n], doc2[n], diff_png)
            #print(f'Page {n}: {diff_png} {sidx_n=}.')
            sidx += sidx_n
        sidx /= n1
        #print(f'{sidx=}')
        return sidx
    finally:
        doc1.close()
        doc2.close()


class TestConversion:
    '''Test the converting process.'''

    def setup_method(self):
        '''create output path if not exist.'''
        if not os.path.exists(output_path): os.mkdir(output_path)


    def convert(self, filename):
        '''Convert PDF file from sample path to output path.'''
        source_pdf_file = os.path.join(sample_path, f'{filename}.pdf')
        docx_file = os.path.join(output_path, f'{filename}.docx')
        c = Converter(source_pdf_file)
        c.convert(docx_file)
        c.close()

    def convert_by_io_stream(self, filename):
        '''Convert PDF file from sample path to output path.'''
        source_pdf_file = os.path.join(sample_path, f'{filename}.pdf')
        with open(source_pdf_file, 'rb') as f: in_stream = f.read()

        c = Converter(stream=in_stream)
        out_stream = io.BytesIO()
        c.convert(out_stream)
        c.close()

        docx_file = os.path.join(output_path, f'{filename}.docx')
        with open(docx_file, 'wb') as f: f.write(out_stream.getvalue())


    # ------------------------------------------
    # table contents
    # ------------------------------------------
    def test_extracting_table(self):
        '''test extracting contents from table.'''
        filename = 'demo-table'
        pdf_file = os.path.join(sample_path, f'{filename}.pdf')
        tables = Converter(pdf_file).extract_tables(end=1)
        print(tables)

        # compare the last table
        table = [[col.strip() if col else col for col in row] for row in tables[-1]]
        sample = [
            ['Input', None, None, None, None, None],
            ['Description A', 'mm', '30.34', '35.30', '19.30', '80.21'],
            ['Description B', '1.00', '5.95', '6.16', '16.48', '48.81'],
            ['Description C', '1.00', '0.98', '0.94', '1.03', '0.32'],
            ['Description D', 'kg', '0.84', '0.53', '0.52', '0.33'],
            ['Description E', '1.00', '0.15', None, None, None],
            ['Description F', '1.00', '0.86', '0.37', '0.78', '0.01']
        ]
        assert table==sample
    
    # ------------------------------------------
    # extract tables with empty/none cells
    # ------------------------------------------
    def test_extract_tables_empty_cell(self):
        '''Test extracting tables from demo-table-empty-cell.pdf (no crash, returns list).'''
        # Test that extract_tables() handles tables with empty cells correctly.
        filename = 'demo-table-empty-cell'
        pdf_file = os.path.join(sample_path, f'{filename}.pdf')
        cv = Converter(pdf_file)
        tables = cv.extract_tables()
        cv.close()
        assert isinstance(tables, list), 'extract_tables() should return a list'
        # At least one table expected from this sample
        assert len(tables) >= 1, f'expected at least 1 table, got {len(tables)}'


    # ------------------------------------------
    # command line arguments
    # ------------------------------------------
    def test_multi_pages(self):
        '''test converting pdf with multi-pages.'''
        filename = 'demo'
        pdf_file = os.path.join(sample_path, f'{filename}.pdf')
        docx_file = os.path.join(output_path, f'{filename}.docx')
        parse(pdf_file, docx_file, start=1, end=5)

        # check file
        assert os.path.isfile(docx_file)

    # ------------------------------------------
    # rotated images (issue 346)
    # ------------------------------------------
    def test_rotated_images(self):
        '''Test converting PDF with rotated images (issue 346).'''
        filename = 'demo-issue-346'
        self.convert(filename)
        docx_file = os.path.join(output_path, f'{filename}.docx')
        assert os.path.isfile(docx_file), f'Expected output file: {docx_file}'
        assert os.path.getsize(docx_file) > 0, 'Output docx should not be empty'
    
    # ------------------------------------------
    # non-grayscale/non-RGB images (issue 340)
    # ------------------------------------------
    def test_non_grayscale_non_rgb_images(self):
        '''Test converting PDFs with rotated images and non-RGB colorspace (CMYK, alpha).

        Covers: per-image rotation, pixmap tobytes() for non-grayscale/non-RGB.
        '''
        filenames = [
            'demo-issue-340'
        ]
        for filename in filenames:
            pdf_file = os.path.join(sample_path, f'{filename}.pdf')
            if not os.path.isfile(pdf_file):
                continue
            docx_file = os.path.join(output_path, f'{filename}.docx')
            parse(pdf_file, docx_file, start=0, end=None)
            assert os.path.isfile(docx_file), f'Expected output {docx_file}'

    def test_scaled_text_matrix_uses_effective_font_size(self):
        '''Test extracting transformed text with effective font size.'''
        content_stream = b'q 0.03 0 0 0.03 72 720 cm BT /F1 488 Tf 0 0 Td (SCALED) Tj ET Q'
        doc = PdfiumDocument(stream=build_pdf_bytes(content_stream))
        page = doc[0]
        try:
            blocks = page.extract_text_blocks()
            sizes = [
                span['size']
                for block in blocks
                for line in block['lines']
                for span in line['spans']
            ]
        finally:
            page.close()
            doc.close()

        assert max(sizes) == pytest.approx(14.64, abs=0.01)

    def test_extracted_images_are_docx_compatible(self):
        '''Test extracted PDF images can be consumed by python-docx.'''
        document = Document()
        for filename in ['demo-text-hidden', 'demo-image-floating']:
            pdf_file = os.path.join(sample_path, f'{filename}.pdf')
            doc = PdfiumDocument(pdf_file)
            page = doc[0]
            try:
                images = page.extract_images()
            finally:
                page.close()
                doc.close()

            assert images
            for image in images:
                document.add_picture(BytesIO(image['image']))

    def test_cropbox_offset_normalizes_text_coordinates(self):
        '''Test text coordinates are normalized to the visible crop box.'''
        pdf_file = os.path.join(sample_path, 'demo-text-scaling.pdf')
        doc = PdfiumDocument(pdf_file)
        page = doc[0]
        try:
            blocks = page.extract_text_blocks()
        finally:
            page.close()
            doc.close()

        first_char = next(
            char
            for block in blocks
            for line in block['lines']
            for span in line['spans']
            for char in span['chars']
            if char['c'].strip()
        )
        assert first_char['bbox'][0] == pytest.approx(66.83, abs=0.01)
        assert first_char['bbox'][1] == pytest.approx(41.18, abs=0.01)

    def test_rendered_region_filters_covered_images(self):
        '''Test standalone images covered by rendered regions are removed.'''
        image_blocks = [
            {'bbox': (10, 10, 20, 20), 'image': b'a'},
            {'bbox': (40, 40, 60, 60), 'image': b'b'},
        ]
        rendered_regions = [
            {'bbox': (0, 0, 30, 30), 'image': b'rendered'},
        ]

        filtered_images = RawPagePdfium._remove_images_covered_by_rendered_regions(
            image_blocks,
            rendered_regions,
        )

        assert filtered_images == [image_blocks[1]]

    def test_borderless_table_cell_preserves_physical_line_breaks(self):
        '''Test borderless table cells preserve source row boundaries.'''
        borderless_cell = build_table_cell_with_two_text_rows()
        text_block = borderless_cell.blocks[0]

        text_block.parse_horizontal_spacing(
            borderless_cell.working_bbox,
            5.0,
            0.5,
            0.1,
            1.0,
            1.0,
            2.0,
        )
        text_block.parse_exact_line_spacing()

        assert [line.line_break for line in text_block.lines] == [1, 0]
        assert text_block.line_space == pytest.approx(24.0)

    def test_bordered_table_cell_keeps_flow_line_breaks(self):
        '''Test bordered table cells keep normal flowing line-break behavior.'''
        bordered_cell = build_table_cell_with_two_text_rows(border_width=(1, 1, 1, 1))
        text_block = bordered_cell.blocks[0]

        text_block.parse_horizontal_spacing(
            bordered_cell.working_bbox,
            5.0,
            0.5,
            0.1,
            1.0,
            1.0,
            2.0,
        )
        text_block.parse_exact_line_spacing()

        assert [line.line_break for line in text_block.lines] == [0, 0]
        assert text_block.line_space == pytest.approx(16.0)

    def test_mixed_float_image_text_block_preserves_text_offset(self):
        '''Test text to the right of a floated image keeps its source x-offset.'''
        document = Document()
        paragraph = document.add_paragraph()
        text_block = build_mixed_image_text_block()

        text_block.make_docx(paragraph)

        assert len(paragraph._element.xpath('.//w:tab')) == 1

    def test_paragraph_format_keeps_schema_order(self):
        '''Test paragraph properties stay in schema order and are not duplicated.'''
        paragraph = Document().add_paragraph()
        paragraph.paragraph_format.tab_stops.add_tab_stop(Pt(72))

        docx_helpers.reset_paragraph_format(paragraph)
        docx_helpers.reset_paragraph_format(paragraph)

        assert_schema_order(paragraph._p.pPr, docx_helpers.PARAGRAPH_PROPERTY_ORDER)
        assert len(paragraph._p.pPr.findall(qn('w:autoSpaceDE'))) == 1

    def test_run_properties_are_schema_valid(self):
        '''Test character scaling and spacing are whole numbers placed after fonts.'''
        run = Document().add_paragraph().add_run('text')
        run.font.name = 'Arial'
        run.font.size = Pt(9)

        docx_helpers.set_char_scaling(run, 0.8333)
        docx_helpers.set_char_spacing(run, 0.37)
        docx_helpers.set_char_underline(run, 0x123456)

        rPr = run._r.rPr
        assert_schema_order(rPr, docx_helpers.RUN_PROPERTY_ORDER)
        assert rPr.find(qn('w:w')).get(qn('w:val')) == '83'
        assert rPr.find(qn('w:spacing')).get(qn('w:val')) == '7'
        assert rPr.find(qn('w:u')).get(qn('w:color')) == '123456'

    def test_cell_properties_are_schema_valid(self):
        '''Test cell borders, shading and margins validate whatever order they are set in.'''
        table = Document().add_table(rows=1, cols=1)
        cell = table.cell(0, 0)

        docx_helpers.set_cell_margins(cell, start=0, end=0)
        docx_helpers.set_cell_shading(cell, 0xD14727)
        docx_helpers.set_cell_border(
            cell,
            start={'sz': 4, 'val': 'single', 'color': 'D14727'},
            bottom={'sz': 4, 'val': 'single', 'color': 'D14727'},
            top={'sz': 4, 'val': 'single', 'color': 'D14727'},
        )
        docx_helpers.indent_table(table, 3.65)

        tcPr = cell._tc.tcPr
        assert_schema_order(tcPr, docx_helpers.CELL_PROPERTY_ORDER)
        assert_schema_order(tcPr.find(qn('w:tcBorders')), docx_helpers.CELL_BORDER_ORDER)
        assert_schema_order(tcPr.find(qn('w:tcMar')), docx_helpers.CELL_MARGIN_ORDER)
        assert tcPr.find(qn('w:shd')).get(qn('w:val')) == 'clear'
        assert_schema_order(table._tbl.tblPr, docx_helpers.TABLE_PROPERTY_ORDER)
        assert table._tbl.tblPr.find(qn('w:tblInd')).get(qn('w:w')) == '73'

    def test_cell_border_values_are_schema_valid(self):
        '''Test fractional PDF border widths become whole eighth-points and colours drop "#".'''
        table = Document().add_table(rows=1, cols=1)
        cell = build_table_cell_with_two_text_rows(border_width=(0.1, 0.5, 0.25, 20))
        cell.border_color = (0x000000, 0xD14727, 0x5B9BD5, 0xFFFFFF)

        cell.make_docx(table, (0, 0))

        borders = table.cell(0, 0)._tc.tcPr.find(qn('w:tcBorders'))
        assert_schema_order(borders, docx_helpers.CELL_BORDER_ORDER)
        assert {
            border.tag: (border.get(qn('w:sz')), border.get(qn('w:color')))
            for border in borders
        } == {
            qn('w:top'): ('2', '000000'),
            qn('w:end'): ('4', 'D14727'),
            qn('w:bottom'): ('2', '5B9BD5'),
            qn('w:start'): ('96', 'FFFFFF'),
        }

    @pytest.mark.parametrize('filename', [
        'demo-table', 'demo-table-border-style', 'demo-table-shading',
        'demo-table-shading-highlight', 'demo-text-scaling', 'demo-section',
    ])
    def test_converted_properties_keep_schema_order(self, filename):
        '''Test every converted property element keeps the order the schema requires.'''
        self.convert(filename)
        document = Document(os.path.join(output_path, f'{filename}.docx'))
        body = document.element.body
        for tag, order in (
            ('w:pPr', docx_helpers.PARAGRAPH_PROPERTY_ORDER),
            ('w:rPr', docx_helpers.RUN_PROPERTY_ORDER),
            ('w:tblPr', docx_helpers.TABLE_PROPERTY_ORDER),
            ('w:tcPr', docx_helpers.CELL_PROPERTY_ORDER),
            ('w:tcBorders', docx_helpers.CELL_BORDER_ORDER),
        ):
            for element in body.iter(qn(tag)):
                assert_schema_order(element, order)
        for border in body.iter(qn('w:tcBorders')):
            for edge in border:
                assert edge.get(qn('w:sz')).isdigit()
                assert not edge.get(qn('w:color')).startswith('#')

    @staticmethod
    def _page_text(blocks):
        return ''.join(
            char['c']
            for block in blocks
            for line in block['lines']
            for span in line['spans']
            for char in span['chars']
        )

    def test_invisible_text_layer_is_dropped(self):
        '''With ocr=0, render-mode-3 (invisible) OCR text behind an image is ignored.'''
        doc = PdfiumDocument(os.path.join(sample_path, 'demo-text-hidden.pdf'))
        page = doc[0]
        try:
            visible = page.extract_text_blocks(include='visible')
        finally:
            page.close()
            doc.close()

        text = self._page_text(visible)
        assert 'displayed text' in text
        assert 'encouraged to read widely' not in text

    def test_ocr_mode_extracts_only_hidden_text(self):
        '''include="hidden" returns only the invisible OCR layer (ocr=2 parity).'''
        doc = PdfiumDocument(os.path.join(sample_path, 'demo-text-hidden.pdf'))
        page = doc[0]
        try:
            hidden = page.extract_text_blocks(include='hidden')
        finally:
            page.close()
            doc.close()

        text = self._page_text(hidden)
        assert 'encouraged to read widely' in text
        assert 'displayed text' not in text

    def test_rotated_page_uses_visual_dimensions(self):
        '''A /Rotate 90 page reports portrait (visible) dimensions and a real rotation matrix.'''
        doc = PdfiumDocument(os.path.join(sample_path, 'demo-image-rotation.pdf'))
        page = doc[0]
        try:
            assert page.rotation == 90
            assert page.width < page.height
            assert bool(page.rotation_matrix)
        finally:
            page.close()
            doc.close()

    def test_native_image_preserves_source_resolution(self):
        '''Images are extracted at native resolution, not capped to a low-DPI raster clip.'''
        doc = PdfiumDocument(os.path.join(sample_path, 'demo-image-colorspace.pdf'))
        page = doc[0]
        try:
            images = page.extract_images()
        finally:
            page.close()
            doc.close()

        assert images
        assert max(image['width'] for image in images) >= 1000

    def test_hyperlinks_are_extracted(self):
        '''URI link annotations are recovered as hyperlink shapes.'''
        doc = PdfiumDocument(os.path.join(sample_path, 'demo-image.pdf'))
        page = doc[0]
        try:
            links = page.extract_hyperlinks()
        finally:
            page.close()
            doc.close()

        assert links
        assert all(link['uri'].startswith('http') for link in links)

    def test_path_with_undefined_color_does_not_crash(self):
        '''A path whose stroke/fill color is undefined falls back to white instead of crashing.'''
        path = Path({
            'type': 'fs',
            'items': [('l', (0.0, 0.0), (10.0, 0.0)), ('l', (10.0, 0.0), (10.0, 10.0)), ('close',)],
            'width': 1.0,
            'color': None,
            'fill': None,
            'closePath': True,
        })

        assert path.to_shapes()

    def test_stroked_regular_font_is_bold(self):
        '''Text emboldened by stroking a regular font in its fill colour is bold.'''
        content_stream = (
            b'BT /F1 12 Tf 72 720 Td (Plain) Tj ET '
            b'0.3 w BT 2 Tr /F1 12 Tf 72 700 Td (Stroked) Tj ET '
            b'0 0 1 RG BT 2 Tr /F1 12 Tf 72 680 Td (Outlined) Tj ET'
        )
        doc = PdfiumDocument(stream=build_pdf_bytes(content_stream))
        page = doc[0]
        try:
            spans = {
                ''.join(char['c'] for char in span['chars']).strip(): span['flags']
                for block in page.extract_text_blocks()
                for line in block['lines']
                for span in line['spans']
            }
        finally:
            page.close()
            doc.close()

        assert not spans['Plain'] & 2**4
        assert spans['Stroked'] & 2**4
        assert not spans['Outlined'] & 2**4

    def test_fill_band_splits_lattice_row(self):
        '''A shaded band without its own ruling becomes its own row, so each row keeps one shading.'''
        rule = lambda x0, y0, x1, y1: f'0 0 0 rg {x0} {y0} {x1 - x0} {y1 - y0} re f '.encode()
        content_stream = (
            b'0.82 0.28 0.15 rg 50 660 500 40 re f '
            b'0.84 0.82 0.80 rg 50 500 500 160 re f '
            + rule(50, 699.6, 550, 700) + rule(50, 500, 550, 500.4)
            + rule(50, 500, 50.4, 700) + rule(549.6, 500, 550, 700)
            + rule(50, 640, 550, 640.4) + rule(300, 500, 300.4, 660)
            + b'1 1 1 rg BT /F1 12 Tf 60 675 Td (HEADER) Tj ET '
            b'0 0 0 rg BT /F1 10 Tf 60 645 Td (LEFT) Tj ET BT /F1 10 Tf 310 645 Td (RIGHT) Tj ET '
            b'BT /F1 10 Tf 60 600 Td (BODY) Tj ET'
        )
        pdf_file = os.path.join(output_path, 'fill-band-table.pdf')
        with open(pdf_file, 'wb') as f:
            f.write(build_pdf_bytes(content_stream))
        cv = Converter(pdf_file)
        try:
            cv.parse(**cv.default_settings)
            table = cv.pages[0].store()['sections'][0]['columns'][0]['blocks'][0]
        finally:
            cv.close()

        rows = table['rows']
        assert [round(row['bbox'][3]) for row in rows] == [132, 152, 292]
        assert rows[0]['cells'][0]['bg_color'] == rgb_value((0.82, 0.28, 0.15))
        assert {cell['bg_color'] for row in rows[1:] for cell in row['cells'] if cell} == {
            rgb_value((0.84, 0.82, 0.80))
        }

    def test_label_after_gap_in_row_reaches_its_tab_stop(self):
        '''A right-hand label starting exactly at a tab stop is positioned by that stop.'''
        def line(x0, x1, text):
            return {'bbox': (x0, 73.5, x1, 83.4), 'wmode': 0, 'dir': (1.0, 0.0), 'spans': [{
                'bbox': (x0, 73.5, x1, 83.4), 'font': 'Arial', 'size': 10.0, 'color': 0, 'flags': 0,
                'chars': [{'c': text, 'bbox': (x0, 73.5, x1, 83.4)}],
            }]}
        # 358.2 - 26.6 is 331.5999... in floating point, just short of the 331.6 stop
        block = TextBlock({'lines': [line(26.6, 137.9, 'TRADE SCOPE'), line(358.2, 396.6, 'Contact:')]})

        block.lines.parse_tab_stop(5.0)

        assert block.tab_stops == [331.6]
        assert [line.tab_stop for line in block.lines] == [0, 1]

    def test_repeated_letterhead_and_page_numbers_move_to_header_and_footer(self):
        '''Furniture repeated across pages becomes the Word header and footer.'''
        pages = [
            text_at(72, 750, 'ACME BUILDERS LETTERHEAD') + text_at(72, 600, body) + text_at(280, 40, f'Page {n} of 2')
            for n, body in ((1, 'First page scope'), (2, 'Second page scope'))
        ]
        document = convert_pages('furniture-repeated', pages)

        assert 'ACME' not in body_text(document) and 'Page 1 of 2' not in body_text(document)
        for section in document.sections:
            assert 'ACME BUILDERS LETTERHEAD' in part_text(section.header)
            assert page_fields(section.footer) == ['PAGE', 'NUMPAGES']
        assert 'First page scope' in body_text(document)

    def test_single_page_footer_needs_a_page_number(self):
        '''A lone page keeps a closing line in the body unless its edge group shows a page number.'''
        plain = convert_pages('furniture-plain', [text_at(72, 700, 'Scope item') + text_at(72, 40, 'Prepared by the estimating team')])
        numbered = convert_pages('furniture-numbered', [text_at(72, 700, 'Scope item') + text_at(500, 40, 'PAGE 1 OF 1')])

        assert 'Prepared by the estimating team' in body_text(plain)
        assert not part_text(plain.sections[0].footer).strip()
        assert 'PAGE 1 OF 1' not in body_text(numbered)
        assert page_fields(numbered.sections[0].footer) == ['PAGE', 'NUMPAGES']

    def test_bundle_page_numbers_stay_text(self):
        '''Numbering that restarts per enclosed document cannot be a document field.'''
        pages = [text_at(72, 700, f'Document {n} scope') + text_at(280, 40, 'Page 1 of 1') for n in (1, 2)]
        document = convert_pages('furniture-bundle', pages)

        for section in document.sections:
            assert 'Page 1 of 1' in part_text(section.footer)
            assert page_fields(section.footer) == []

    def test_footer_rows_stay_separate_paragraphs(self):
        '''A right-hand page number above a left-hand date keeps both positions.'''
        content = text_at(72, 700, 'Scope item') + text_at(500, 40, 'PAGE 1 OF 1', 6) + text_at(55, 30, 'PRINTED: 6/7/2023', 6)
        footer = convert_pages('furniture-rows', [content]).sections[0].footer

        texts = [''.join(t.text or '' for t in p._p.iter(qn('w:t'))) for p in footer.paragraphs]
        assert [t for t in texts if t.strip()] == ['PAGE 1 OF 1', 'PRINTED: 6/7/2023']

    def test_full_pages_do_not_spill_blank_pages(self):
        '''The paragraph holding each page's section break needs no room of its own.'''
        pages = [text_at(72, 750, 'LETTERHEAD') + text_at(72, 600, f'Scope {n}') + text_at(280, 40, f'Page {n} of 3') for n in (1, 2, 3)]
        document = convert_pages('furniture-breaks', pages)

        # titles that differ only by a number are body text, not a repeated header
        assert all(f'Scope {n}' in body_text(document) for n in (1, 2, 3))
        breaks = [p for p in document.paragraphs if p._p.find(f"{qn('w:pPr')}/{qn('w:sectPr')}") is not None]
        assert len(breaks) == 2
        for break_paragraph in breaks:
            assert break_paragraph._p.pPr.find(qn('w:spacing')).get(qn('w:lineRule')) == 'exact'
            previous = break_paragraph._p.getprevious()
            assert previous.tag != qn('w:p') or previous.find(qn('w:r')) is not None

    def test_empty_cell_takes_row_text_spacing(self):
        '''A bare cell copies the row text's spacing and size; a cell holding a run keeps its own.'''
        table = Document().add_table(rows=1, cols=3)
        text, empty, image = table.rows[0].cells
        paragraph = text.paragraphs[0]
        docx_helpers.reset_paragraph_format(paragraph, Pt(8.5))
        paragraph.add_run('Exclude sales tax').font.size = Pt(8.5)
        image.paragraphs[0].add_run()

        Row._match_empty_cells_to_row_text(table.rows[0])

        spacing = empty.paragraphs[0]._p.pPr.find(qn('w:spacing'))
        assert (spacing.get(qn('w:line')), spacing.get(qn('w:lineRule')), spacing.get(qn('w:after'))) == ('170', 'exact', '0')
        assert empty.paragraphs[0]._p.pPr.find(f"{qn('w:rPr')}/{qn('w:sz')}").get(qn('w:val')) == '17'
        assert_schema_order(empty.paragraphs[0]._p.pPr, docx_helpers.PARAGRAPH_PROPERTY_ORDER)
        assert image.paragraphs[0]._p.pPr is None

    def test_text_after_list_marker_gap_is_not_tabbed(self):
        '''A gap narrower than a column (e.g. after "a.") stays a space, so flowing text keeps its wrap.'''
        def line(x0, x1, text):
            return {'bbox': (x0, 73.5, x1, 83.4), 'wmode': 0, 'dir': (1.0, 0.0), 'spans': [{
                'bbox': (x0, 73.5, x1, 83.4), 'font': 'Arial', 'size': 10.0, 'color': 0, 'flags': 0,
                'chars': [{'c': text, 'bbox': (x0, 73.5, x1, 83.4)}],
            }]}
        # 41.8 - 26.6 is 15.1999... in floating point, just short of the 15.2 stop
        block = TextBlock({'lines': [line(26.6, 36.6, 'a.'), line(41.8, 300.0, 'Eastern and southern')]})

        block.lines.parse_tab_stop(5.0)

        assert block.tab_stops == [15.2]
        assert [line.tab_stop for line in block.lines] == [0, 0]


# We make a separate pytest test for each sample file.

def _find_paths():
    ret = list()
    for path in glob.glob(f'{sample_path}/*.docx') + glob.glob(f'{sample_path}/*.pdf'):
        path_leaf = os.path.basename(path)
        if path_leaf.count('.') > 1:
            continue
        if path_leaf == 'demo-whisper_2_3.pdf':
            # Known to fail.
            continue
        ret.append(os.path.relpath(path, root_path))
    return ret

g_paths = _find_paths()

# We create a separate pytest for each sample file, paramaterised using the
# path of the sample file relative to the repository root.
#
# So one can run a specific test with:
#
# pytest test/test.py::test_one[test/samples/demo-whisper_2_3.pdf]

@pytest.mark.parametrize('path', g_paths)
def test_one(path):
    '''Check the quality of converted docx.
    '''
    
    # Where there are two values, they are (sidx_required_word,
    # sidx_required_libreoffice).
    #
    docx_to_sidx_required = {
        'demo-blank.pdf': 1.0,
        'demo-image-cmyk.pdf': 0.90,
        'demo-image-transparent.pdf': 0.90,
        'demo-image-vector-graphic.pdf': (0.89, 0.64),
        'demo-image.pdf': 0.90,
        'demo-image-rotation.pdf': (0.90, 0.82),
        'demo-image-overlap.pdf': (0.90, 0.70),
        'demo-path-transformation.pdf': (0.89, 0.60),
        'demo-section-spacing.pdf': (0.90, 0.86),
        'demo-section.pdf': (0.70, 0.45),
        'demo-table-align-borders.pdf': 0.49,
        'demo-table-border-style.pdf': (0.90, 0.89),
        'demo-table-bottom.pdf': 0.90,
        'demo-table-close-underline.pdf': (0.57, 0.49),
        'demo-table-lattice-one-cell.pdf': (0.79, 0.75),
        'demo-table-lattice.pdf': (0.75, 0.59),
        'demo-table-nested.pdf': 0.84,
        'demo-table-shading-highlight.pdf': (0.55, 0.45),
        'demo-table-shading.pdf': (0.80, 0.59),
        'demo-table-stream.pdf': 0.55,
        'demo-table.pdf': (0.90, 0.75),
        'demo-text-alignment.pdf': (0.90, 0.86),
        'demo-text-scaling.pdf': (0.80, 0.62),
        'demo-text-unnamed-fonts.pdf': (0.80, 0.77),
        'demo-text-hidden.pdf': 0.90,
        'demo-text.pdf': 0.75,
        'docx-from-pdf-lists-bullets3.docx': (0.98, 0.99),
    }

    print(f'# Looking at: {path}')
    path = f'{root_path}/{path}'
    path_leaf = os.path.basename(path)
    _, ext = os.path.splitext(path)
    if ext == '.docx':
        pdf = f'{path}.pdf'
        document_to(path, pdf)
    else:
        pdf = path
    docx2 = f'{pdf}.docx'
    pages = None
    if os.path.basename(path) == 'demo-whisper_2_3.pdf':
        pages = [25, 26, 27]
    else:
        doc = pdfium.PdfDocument(pdf)
        try:
            if len(doc) > 1:
                print(f'Not testing because more than one page: {path}')
                return
        finally:
            doc.close()
    #print(f'Calling parse() {pdf=} {docx2=}')
    parse(pdf, docx2, pages=pages, raw_exceptions=True)
    assert os.path.isfile(docx2)
    pdf2 = f'{docx2}.pdf'
    document_to(docx2, pdf2)
    sidx = compare_pdf(pdf, pdf2, num_pages=1)

    sidx_required = docx_to_sidx_required.get(path_leaf)
    if sidx_required:
        if isinstance(sidx_required, tuple):
            sr_word, sr_libreoffice = sidx_required
            sidx_required = sr_word if platform.system() == 'Windows' else sr_libreoffice

        #print(f'{path=}: {sidx_required=} {sidx=}.')
        comparison_epsilon = 1e-6
        if sidx + comparison_epsilon < sidx_required:
            print(f'{sidx=} too low - should be >= {sidx_required=}')
            print(f'    {pdf}')
            print(f'    {pdf2}')
            assert 0
    else:
        print(f'# No sidx_required available for {path_leaf=}.')
