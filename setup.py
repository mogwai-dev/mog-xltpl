import os
from io import open
from setuptools import setup

CUR_DIR = os.path.abspath(os.path.dirname(__file__))
README = os.path.join(CUR_DIR, "README_EN.md")
with open(README, 'r', encoding='utf-8') as fd:
    long_description = fd.read()

setup(
    name = 'mog-xltpl',
    version = "1.0.0",
    author = 'mog (forked from Zhang Yu)',
    author_email = '',
    url = 'https://github.com/[YOUR-USERNAME]/mog-xltpl',
    packages = ['xltpl'],
	install_requires = [
        'xlrd >= 1.2.0',
        'xlwt >= 1.3.0',
        'openpyxl >= 3.1.0',
        'jinja2',
        'six',
        'pywin32 >= 311',
        'pillow >= 9.0.0',
        'pyyaml >= 6.0.3',
        'tomli'
    ],
    description = ( 'A Windows-only xltpl fork that preserves VBA, images, and complex Excel formatting using COM' ),
    long_description = long_description,
    long_description_content_type = "text/markdown",
    platforms = ["Windows"],
    license = 'MIT',
    keywords = ['Excel', 'xls', 'xlsx', 'spreadsheet', 'workbook', 'template', 'VBA', 'COM', 'Windows', 'preserve', 'formatting'],
    classifiers=[
        'Development Status :: 4 - Beta',
        'Intended Audience :: Developers',
        'License :: OSI Approved :: MIT License',
        'Operating System :: Microsoft :: Windows',
        'Programming Language :: Python :: 3',
        'Programming Language :: Python :: 3.8',
        'Programming Language :: Python :: 3.9',
        'Programming Language :: Python :: 3.10',
        'Programming Language :: Python :: 3.11',
        'Programming Language :: Python :: 3.12',
        'Programming Language :: Python :: 3.13',
        'Topic :: Office/Business :: Financial :: Spreadsheet',
    ],
)
