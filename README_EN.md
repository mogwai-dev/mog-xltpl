
# mog-xltpl

[![PyPI version](https://badge.fury.io/py/mog-xltpl.svg)](https://pypi.org/project/mog-xltpl/)
[![Python 3.8+](https://img.shields.io/badge/python-3.8+-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

A Windows-only **CLI tool** and Python module to generate `.xlsx`/`.xlsm` files from templates while preserving VBA, images, and complex Excel formatting using COM. [中文](README.md) | [日本語](README_JA.md)

**Primary Use**: Command-line tool for templated Excel document generation via Taskfile integration.

> **Note**: This tool is designed to work with [Taskfile](https://taskfile.dev/). It uses YAML `vars:` sections for variable definitions, allowing you to specify templates and output files from Taskfile.
> Taskfile-style `{{ .VAR }}` is not a full Go template implementation; we only normalize it to `{{ VAR }}` before rendering.

## How it works

When xltpl reads a `.xlsx`/`.xlsm` file, it creates a tree for each worksheet.  
Each tree is translated to a Jinja2 template with custom tags.  
When the template is rendered, Jinja2 extensions of custom tags call corresponding tree nodes to write the output file.

**Key Feature**: Uses Excel COM API (via `pywin32`) to ensure complete preservation of images, drawings, and Excel-specific formatting that other tools cannot maintain.

## How to install

```shell
pip install mog-xltpl
```

**or with uv tool (recommended):**

```shell
uv tool install mog-xltpl
```

**Requirements:**
- Windows OS
- Python 3.8+
- Microsoft Excel (COM API access required)
- `pywin32` (automatically installed, required for preserving images and formatting)

> **Note**: This tool requires Windows and Microsoft Excel because it uses Excel COM API to ensure complete preservation of images, drawings, and Excel-specific formatting. Without Excel or on non-Windows systems, the tool will exit with an error message.

### Develop & test with uv

```shell
uv venv
uv pip install -e .[test]
uv run pytest
```

## Quick Start: CLI Usage (Recommended)

### Simple Usage

Specify template file, output file, and variables file:

```shell
mog-xltpl template.xlsx output.xlsx vars.yaml

# To emit an additional highlighted copy (auto-named as output_highlight.xlsx)
mog-xltpl template.xlsx output.xlsx vars.yaml --highlight-output

# If you want to set color explicitly
mog-xltpl template.xlsx output.xlsx vars.yaml \
  --highlight-output \
  --highlight-color FFFF9999
```

### Integration with Taskfile (Recommended Workflow)

**Taskfile.yml:**

```yaml
version: '3'

vars:
  DOC_TYPE: invoice
  DATE: "2025-12-30"
  NAME: "John Doe"

tasks:
  render:
    cmds:
      - mog-xltpl templates/{{.DOC_TYPE}}.xlsx output/result.xlsx vars.yaml
```

**vars.yaml:**

```yaml
vars:
  doc_type: "invoice"
  date: "2025-12-30"
  name: "John Doe"
  items:
    - name: "Product A"
      price: 1000
    - name: "Product B"
      price: 2000
```

**Run:**

```bash
task render
```

### Path Expansion Rules
- Template and output files are specified from the command line.
- YAML file contains only the `vars` section.
- Relative paths are resolved from the execution directory.
- Use `--highlight-output` to auto-emit a highlighted copy named `<output>_highlight` (color via `--highlight-color`, e.g., `FFFF9999`).

### Vars Resolution
- `vars` accepts either a mapping or a list of single-key mappings (Taskfile style: `- KEY: value`).
- Values are rendered against the same `vars` map for a few passes, so self-references like `FOOBAR: "foo_{{ VAR }}"` are expanded before the workbook is rendered.

## Python API (Advanced Usage)

*   To use xltpl, you need to be familiar with the [syntax of jinja2 template](https://jinja.palletsprojects.com/).
*   Get a pre-written xls/x file as the template.
*   Insert variables in the cells, such as : 

```jinja2
{{name}}
```
  
*   ~~Insert control statements in the notes(comments) of cells, use beforerow, beforecell or aftercell to seperate them :~~


```jinja2
beforerow{% for item in items %}
```
```jinja2
beforerow{% endfor %}
```

*   Insert control statements in the cells (**v0.9**) :

```jinja2
{%- for row in rows %}
{% set outer_loop = loop %}{% for row in rows %}
Cell
{{outer_loop.index}}{{loop.index}}
{%+ endfor%}{%+ endfor%}
```

**Image insertion**

To insert images, use the `img` filter:

```jinja2
{{ image_path | img(120, 140) }}
```

- First argument: Path to image file
- Second argument (optional): Width in pixels
- Third argument (optional): Height in pixels

You can also use keyword arguments:

```jinja2
{{ image_path | img(width=120, height=140) }}
```

**Other handy filters**

- `sha256`: `{{ file | sha256 }}`
- `mtime`: `{{ file | mtime('%Y-%m-%d') }}`
- `to_fullwidth`: convert half-width digits and `-` to full-width for Excel-friendly formatting

*   Run the code
```python
from xltpl.writerx import BookWriter
writer = BookWriter('tpl.xlsx')
person_info = {'name': u'Hello Wizard'}
items = ['1', '1', '1', '1', '1', '1', '1', '1', ]
person_info['items'] = items
payloads = [person_info]
writer.render_book(payloads)
writer.save('result.xlsx')
```

## Supported
* xls (xlrd/xlwt) / xlsx and xlsm (openpyxl)
* MergedCell   
* Non-string value for a cell (use **{% xv variable %}** to specify a variable) 
* For xlsx family  
Image (use **{% img variable %}**)  
DataValidation   
AutoFilter

## Architecture and Image Preservation

### Why pywin32/COM API?

This tool uses **Excel COM API** (via pywin32) for saving files to ensure complete preservation of:
- **Images and drawings** embedded in templates
- **All Excel namespaces** and XML attributes
- **Complex formatting** and workbook properties
- **Macro-enabled files** (.xlsm)

**Previous approach using openpyxl had limitations:**
- openpyxl removes images and drawings when saving
- openpyxl strips Excel-specific XML namespaces
- Result files often couldn't be opened by Excel

**Current approach (pywin32 + COM API):**
1. Load template using openpyxl (read-only, for Jinja2 rendering)
2. Open template copy using Excel COM API
3. Update only cell values from rendered data
4. Save via COM API → All images, drawings, and formatting preserved

### Testing Reproducibility

To verify that images and formatting are preserved:

```bash
# Create a test template with images
# Use static_image.xlsm as example

# Run rendering
xltpl static_image.xlsm static_image_out.xlsm static_image.yaml

# Verify file integrity
python -c "
import zipfile
z = zipfile.ZipFile('static_image_out.xlsm')
images = [n for n in z.namelist() if 'image' in n.lower() or 'drawing' in n.lower()]
print(f'Images/drawings preserved: {len(images)}')
for img in images:
    print(f'  {img}')
"

# Compare with template
python -c "
import zipfile
t = zipfile.ZipFile('static_image.xlsm')
o = zipfile.ZipFile('static_image_out.xlsm')
t_imgs = set([n for n in t.namelist() if 'drawing' in n or 'media' in n])
o_imgs = set([n for n in o.namelist() if 'drawing' in n or 'media' in n])
print(f'Template: {len(t_imgs)} image-related files')
print(f'Output: {len(o_imgs)} image-related files')
print(f'Match: {t_imgs == o_imgs}')
"

# Verify namespaces preserved
python -c "
import zipfile
o = zipfile.ZipFile('static_image_out.xlsm')
sheet_xml = o.read('xl/worksheets/sheet1.xml').decode('utf-8')
preserved = all([
    'xmlns:mc' in sheet_xml,
    'mc:Ignorable' in sheet_xml,
    'xr:uid' in sheet_xml
])
print(f'Excel namespaces preserved: {preserved}')
"
```

### Parallel Execution Safety

The tool uses a threading lock to prevent concurrent Excel COM operations:
- Multiple xltpl processes can run in parallel (e.g., via Taskfile)
- Each process acquires a lock before using Excel COM
- Prevents COM conflicts and ensures stability

Example with Taskfile parallel execution:

```yaml
tasks:
  process-all:
    deps:
      - task: process-file-1  # Runs in parallel
      - task: process-file-2  # Runs in parallel
    cmds:
      - echo "All files processed"
```

## Related
* [pydocxtpl](https://github.com/zhangyu836/pydocxtpl)  
A python module to generate docx files from a docx template.
* [django-excel-export](https://github.com/zhangyu836/django-excel-export)  
A Django library for exporting data in xlsx, xls, docx format, utilizing xltpl and pydocxtpl, with admin integration.  
[Demo project](https://github.com/zhangyu836/django-excel-export-demo)   
[Live demo](https://tranquil-tundra-83829.herokuapp.com/) (User name: admin
Password: admin)   

* [xltpl for nodejs](https://github.com/zhangyu836/node-xlsx-template)   
CodeSandbox examples: 
[browser](https://codesandbox.io/s/xlsx-export-with-exceljs-and-xltpl-58j9g6)
[node](https://codesandbox.io/s/exceljs-template-with-xltpl-4w58xo)    
* [xltpl for java](https://github.com/zhangyu836/xltpl4java)


## Notes

### xlrd

xlrd does not extract print settings.   
[This repo](https://github.com/zhangyu836/xlrd) does. 

### xlwt
  
xlwt always sets the default font to 'Arial'.  
Excel measures column width units based on the default font.   
[This repo](https://github.com/zhangyu836/xlwt) does not.  
