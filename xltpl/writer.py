# -*- coding: utf-8 -*-

import six

from .base import SheetBase, BookBase
from .writermixin import SheetMixin, BookMixin, Box
from .utils import tag_test, parse_cell_tag
from .xlnode import Tree, Row, Cell, EmptyCell, Node, create_cell
from .jinja import JinjaEnv
from .nodemap import NodeMap
from .sheetresource import SheetResourceMap
from .richtexthandler import rich_handler
from .merger import Merger
from .config import config
from .celltag import CellTag
from pathlib import Path

class SheetWriter(SheetBase, SheetMixin):

    def __init__(self, bookwriter, sheet_resource, sheet_name):
        self.rdbook = bookwriter.rdbook
        self.wtbook = bookwriter.wtbook
        self.style_list = bookwriter.style_list
        self.merger = sheet_resource.merger
        self.rdsheet = sheet_resource.rdsheet
        self.highlight = getattr(bookwriter, 'highlight', False)
        self.highlight_color = getattr(bookwriter, 'highlight_color', None)
        self.log_path = getattr(bookwriter, 'log_path', None)
        self.create_worksheet(self.rdsheet, sheet_name)
        self.wtrows = set()
        self.wtcols = set()
        self.box = Box(-1, -1)

class BookWriter(BookBase, BookMixin):
    sheet_writer_cls = SheetWriter

    def __init__(self, fname, debug=False, highlight=False, highlight_color=None, log_path=None):
        config.debug = debug
        self.debug = debug
        self.highlight = highlight
        self.highlight_color = highlight_color
        self.log_path = log_path
        self.template_path = Path(fname)
        self.load(fname)

    def load(self, fname):
        self.workbook = self.load_rdbook(fname)
        self.font_map = {}
        self.node_map = NodeMap()
        self.jinja_env = JinjaEnv(self.node_map)
        self.merger_cls = Merger
        self.sheet_writer_map = {}
        # Keep track of template sheets to preserve unprocessed ones later
        self.template_sheets = {}  # name -> sheet mapping
        self.processed_sheet_names = set()
        self.sheet_resource_map = SheetResourceMap(self, self.jinja_env)
        for index, rdsheet in enumerate(self.rdbook.sheets()):
            self.template_sheets[rdsheet.name] = rdsheet
            self.sheet_resource_map.add(rdsheet, rdsheet.name, index)

    def build(self, sheet, index, merger):
        tree = Tree(index, self.node_map)
        for rowx in range(sheet.nrows):
            row_node = Row(rowx)
            tree.add_child(row_node)
            for colx in range(sheet.ncols):
                try:
                    sheet_cell = sheet.cell(rowx, colx)
                except Exception as e:
                    raise RuntimeError(f"Failed to read cell at ({rowx}, {colx}): {e}")
                cell_tag_map = None
                note = sheet.cell_note_map.get((rowx, colx))
                if note:
                    comment = note.text
                    if tag_test(comment):
                        _, cell_tag_map = parse_cell_tag(comment)
                value = sheet_cell.value
                cty = sheet_cell.ctype
                rich_text = self.get_rich_text(sheet, rowx, colx)
                if isinstance(value, six.text_type):
                    if not tag_test(value):
                        if rich_text:
                            cell_node = Cell(sheet_cell, rowx, colx, rich_text, cty)
                        else:
                            cell_node = Cell(sheet_cell, rowx, colx, value, cty)
                    else:
                        font = self.get_font(sheet, rowx, colx)
                        cell_node = create_cell(sheet_cell, rowx, colx, value, rich_text, cty, font, rich_handler)
                else:
                    cell_node = Cell(sheet_cell, rowx, colx, value, cty)
                if cell_tag_map:
                    cell_tag = CellTag(cell_tag_map)
                    cell_node.extend_cell_tag(cell_tag)
                    if colx==0:
                        row_node.cell_tag = cell_tag
                tree.add_child(cell_node)
        tree.add_child(Node())#
        return tree

    def render_sheet(self, payload, left_top=None):
        if not hasattr(self, 'wtbook') or self.wtbook is None:
            self.create_workbook()
        return BookMixin.render_sheet(self, payload, left_top)

    def save(self, fname):
        if self.wtbook is not None:
            # Copy unprocessed sheets from template
            if hasattr(self, 'template_sheets') and hasattr(self, 'processed_sheet_names'):
                for sheet_name, rdsheet in self.template_sheets.items():
                    # Copy sheets that haven't been processed (rendered)
                    if sheet_name not in self.processed_sheet_names:
                        # For xls format, we can use xlwt to copy sheets
                        # Since xlwt is used for writing, we use xlrd for reading template
                        # We'll copy cells manually
                        new_sheet = self.wtbook.add_sheet(sheet_name)
                        # Copy all cells from template sheet
                        for rowx in range(rdsheet.nrows):
                            for colx in range(rdsheet.ncols):
                                cell = rdsheet.cell(rowx, colx)
                                if cell.value is not None:
                                    try:
                                        new_sheet.write(rowx, colx, cell.value)
                                    except Exception as e:
                                        # Surface a clear error if copying a cell fails so templates are easier to debug
                                        raise RuntimeError(
                                            f"Failed to write cell at ({rowx}, {colx}) in sheet '{sheet_name}': {e}"
                                        )

            stream = open(fname, 'wb')
            self.wtbook.save(stream)
            stream.close()
            del self.wtbook
        self.sheet_writer_map.clear()
