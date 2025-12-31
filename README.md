# mog-xltpl
A python module to generate xlsx/xlsm files from a xlsx/xlsm template. [中文](README_ZH.md) | [日本語](README_JA.md)

> **Note**: This tool is designed to work with [Taskfile](https://taskfile.dev/). It uses YAML `vars:` sections for variable definitions, allowing you to specify templates and output files from Taskfile.
> Taskfile-style `{{ .VAR }}` is not a full Go template implementation; we only normalize it to `{{ VAR }}` before rendering.

> **Note**: `.xls` (BIFF8) templates are not supported. Use `.xlsx`, `.xlsm`, `.xltx`, or `.xltm` instead. The CLI will exit with an error if a `.xls` file is passed.

## How it works

When xltpl reads a xls/x file, it creates a tree for each worksheet.  
And, each tree is translated to a jinja2 template with custom tags.  
When the template is rendered, jinja2 extensions of custom tags call corresponding tree nodes to write the xls/x file.

## How to install

```shell
pip install xltpl
```

**Requirements:**
- Python 3.8+
- pywin32 (Windows only, required for preserving images and formatting)
- Template/output formats: `.xlsx`, `.xlsm`, `.xltx`, `.xltm` (no `.xls`)

> **Note**: This tool uses Excel COM API via pywin32 to ensure complete preservation of images, drawings, and Excel-specific formatting. Without pywin32, the tool will fail with a clear error message.

### Develop & test with uv

```shell
uv venv
uv pip install -e .[test]
uv run pytest
```

### CLI (template + YAML)

Specify template file, output file, and variables file:

```shell
uv run xltpl template.xlsx output.xlsx vars.yaml

# To emit an additional highlighted copy (auto-named as output_highlight.xlsx)
uv run xltpl template.xlsx output.xlsx vars.yaml --highlight-output

# If you want to set color explicitly
uv run xltpl template.xlsx output.xlsx vars.yaml \
  --highlight-output \
  --highlight-color FFFF9999
```

**Integration with Taskfile:**

```yaml
# Taskfile.yml
version: '3'

vars:
  DOC_TYPE: invoice
  DATE: "2025-12-30"
  NAME: "John Doe"

tasks:
  render:
    cmds:
      - xltpl templates/{{.DOC_TYPE}}.xlsx output/result.xlsx vars.yaml
```

YAML file contains only variables (same style as Taskfile's `vars:`):

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

#### Path Expansion Rules
- Template and output files are specified from the command line.
- YAML file contains only the `vars` section.
- Relative paths are resolved from the execution directory.
- Use `--highlight-output` to auto-emit a highlighted copy named `<output>_highlight` (color via `--highlight-color`, e.g., `FFFF9999`).

#### Vars resolution
- `vars` accepts either a mapping or a list of single-key mappings (Taskfile style: `- KEY: value`).
- Values are rendered against the same `vars` map for a few passes, so self-references like `FOOBAR: "foo_{{ .BAR }}"` are expanded before the workbook is rendered.

## How to use

*   To use xltpl, you need to be familiar with the [syntax of jinja2 template](https://jinja.palletsprojects.com/).
*   Get a pre-written xls/x file as the template.
*   Insert variables in the cells, such as : 

```jinja2
{{name}}
```
  
*   ~~Insert control statements in the notes(comments) of cells, use beforerow, beforecell or aftercell to separate them :~~


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

## Supported / Not supported
- Supported: `.xlsx`, `.xlsm`, `.xltx`, `.xltm`
- Not supported: `.xls` (BIFF8). Convert to `.xlsx` before use.
- Features: merged cells, data validation, autofilter, images (`{% img %}`), non-string cell values (`{% xv %}`).

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
