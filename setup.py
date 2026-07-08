'''Build package.'''

import os
from setuptools import find_packages, setup

DESCRIPTION = 'Python library and CLI for converting PDF files to DOCX.'
EXCLUDE_FROM_PACKAGES = ["build", "dist", "test"]


def get_version(fname):
    '''Read version number from version.txt, created dynamically per Github Action.'''
    if os.path.exists(fname):
        with open(fname, "r", encoding="utf-8") as f:
            version = f.readline().strip()
    else:
        version = '0.5.6a1'
    return version

def load_long_description(fname):
    '''Load README.md for long description'''
    if os.path.exists(fname):
        with open(fname, "r", encoding="utf-8") as f:
            long_description = f.read()
    else:
        long_description = DESCRIPTION

    return long_description

def load_requirements(fname):
    '''Load requirements.'''
    ret = list()
    with open(fname) as f:
        for line in f:
            ret.append(line)
    return ret


setup(
    name="docx-from-pdf",
    version=get_version("version.txt"),
    keywords=["pdf-to-word", "pdf-to-docx", "docx-from-pdf"],
    description=DESCRIPTION,
    long_description=load_long_description("README.md"),
    long_description_content_type="text/markdown",
    license="MIT",
    author='Michael Ellis',
    url='https://github.com/michael-provision/docx-from-pdf/',
    packages=find_packages(exclude=EXCLUDE_FROM_PACKAGES),
    include_package_data=True,
    zip_safe=False,
    install_requires=load_requirements("requirements.txt"),
    python_requires=">=3.10",
    entry_points={
        "console_scripts": [
            "docx-from-pdf=docx_from_pdf.main:main"
            ]},
)
