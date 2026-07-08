'''
Load :py:class:`~docx_from_pdf.page.RawPage` with specified pdf engine.
'''

from .RawPagePdfium import RawPagePdfium


class RawPageFactory:

    MAP = {
        "PDFIUM": RawPagePdfium,
    }

    @classmethod
    def create(cls, page_engine, backend: str = "pdfium"):
        '''Create RawPage class with specified backend.'''
        klass = cls.MAP.get(backend.upper(), None)
        if not klass:
            raise TypeError(f'Page with pdf engine "{backend}" is not implemented yet.')
        else:
            return klass(page_engine=page_engine)
