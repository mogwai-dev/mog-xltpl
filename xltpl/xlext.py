# -*- coding: utf-8 -*-

import os
import six
from jinja2 import nodes
from jinja2.ext import Extension
from jinja2.runtime import Undefined
from inspect import isfunction

class NodeExtension(Extension):
    tags = set(['row', 'cell', 'node', 'extra'])

    def parse(self, parser):
        lineno = next(parser.stream).lineno
        args = [parser.parse_expression()]
        body = []
        return nodes.CallBlock(self.call_method('_node', args),
                               [], [], body).set_lineno(lineno)

    def _node(self, key, caller):
        node = self.environment.node_map.get_node(key)
        return str(key)

class SegmentExtension(Extension):
    tags = set(['seg'])

    def parse(self, parser):
        lineno = next(parser.stream).lineno
        args = [parser.parse_expression()]
        body = parser.parse_statements(['name:endseg'], drop_needle=True)
        return nodes.CallBlock(self.call_method('_seg', args),
                               [], [], body).set_lineno(lineno)

    def _seg(self, key, caller):
        segment = self.environment.node_map.get_node(key)
        rv = caller()
        rv = segment.process_rv(rv)
        return rv

class XvExtension(Extension):
    tags = set(['xv'])

    def parse(self, parser):
        lineno = next(parser.stream).lineno
        args = [parser.parse_expression()]
        if parser.stream.skip_if('comma'):
            args.append(parser.parse_expression())
        else:
            args.append(nodes.Const(0))
        body = []
        return nodes.CallBlock(self.call_method('_xv', args),
                               [], [], body).set_lineno(lineno)

    def _xv(self, xv, key, caller):
        if key==0:
            return six.text_type(xv)
        xvcell = self.environment.node_map.get_node(key)
        if xv is None or type(xv) is Undefined:
            xv = ''
        
        # Check if xv is an ImageMarker (from {{ image_path | img() }} syntax)
        from .filters import ImageMarker
        if isinstance(xv, ImageMarker):
            # Convert to ImageRef and set on cell
            image_ref = ImageRef(xv.path, 0, xv.width, xv.height)
            if not hasattr(xvcell, 'image_ref'):
                xvcell.image_ref = image_ref
            # Set cell value to empty
            xv = ''
        
        xvcell.rv = xv
        
        # Check if there are any pending image_refs from filters (legacy {% xv %} support)
        # The img_filter may have stored image_ref during xv evaluation
        if hasattr(self.environment, '_pending_image_ref') and self.environment._pending_image_ref:
            pending_ref = self.environment._pending_image_ref
            # Set it on the current cell (xvcell)
            if not hasattr(xvcell, 'image_ref'):
                xvcell.image_ref = pending_ref
            # Clear the pending ref
            self.environment._pending_image_ref = None
        
        return six.text_type(xv)

class OpExtension(Extension):
    tags = set(['op'])

    def parse(self, parser):
        lineno = next(parser.stream).lineno
        args = [parser.parse_expression()]
        func_args = []
        while parser.stream.skip_if('comma'):
            func_args.append(parser.parse_expression())
        args.append(nodes.List(func_args))
        body = []
        return nodes.CallBlock(self.call_method('_op', args),
                               [], [], body).set_lineno(lineno)

    def _op(self, func, func_args, caller):
        if(isfunction(func)):
            node = self.environment.node_map.current_node
            node.add_op((func, func_args))
        return six.text_type(func)

class NoopExtension(Extension):
    tags = set(['op'])

    def parse(self, parser):
        lineno = next(parser.stream).lineno
        args = [parser.parse_expression()]
        func_args = []
        while parser.stream.skip_if('comma'):
            func_args.append(parser.parse_expression())
        args.append(nodes.List(func_args))
        body = []
        return nodes.CallBlock(self.call_method('_op', args),
                               [], [], body).set_lineno(lineno)

    def _op(self, func, func_args, caller):
        return six.text_type(func)

class ImageExtension(Extension):
    tags = set(['img'])

    def parse(self, parser):
        lineno = next(parser.stream).lineno
        # Parse image path and optional dimensions using same pattern as OpExtension
        image_path = parser.parse_expression()
        dimensions = []
        while parser.stream.skip_if('comma'):
            dimensions.append(parser.parse_expression())
        
        args = [image_path, nodes.List(dimensions)]
        body = []
        return nodes.CallBlock(self.call_method('_image', args),
                               [], [], body).set_lineno(lineno)

    def _image(self, image_path, dimensions, caller):
        """Insert an image at the current cell position.
        
        Args:
            image_path: Path to the image file
            dimensions: List of [width, height] (optional)
        """
        if not image_path:
            return ''
        # Extract width and height from dimensions list
        width = dimensions[0] if len(dimensions) > 0 else None
        height = dimensions[1] if len(dimensions) > 1 else None
        
        
        # Create ImageRef with the provided path and dimensions
        image_ref = ImageRef(str(image_path), 0, width, height)
        
        # Get current cell node from node_map
        node = self.environment.node_map.current_node
        if hasattr(node, 'current_cell'):
            cell = node.current_cell
            # Set image_ref on the cell so it gets processed during rendering
            if not hasattr(cell, 'image_ref'):
                cell.image_ref = image_ref
        
        return ''

try:
    pil = True
    from PIL.ImageFile import ImageFile
except ImportError as e:
    print(f"Warning: PIL/Pillow import failed, image processing disabled: {e}")
    pil = False

class ImageRef():

    def __init__(self, image, image_index, width=None, height=None):
        self.image = image
        self.image_index = image_index
        self.rdrowx = -1
        self.rdcolx = -1
        self.wtrowx = -1
        self.wtcolx = -1
        self.width = width
        self.height = height
        if not isinstance(image, ImageFile):
            fname = six.text_type(image)
            if not os.path.exists(fname):
                self.image = None

    @property
    def image_key(self):
        return (self.rdrowx,self.rdcolx,self.image_index)

    @property
    def wt_top_left(self):
        return (self.wtrowx,self.wtcolx)

class ImagexExtension(Extension):
    tags = set(['img'])

    def parse(self, parser):
        lineno = next(parser.stream).lineno
        # Parse image path and optional dimensions
        # Support both comma-separated: {% img path, width, height %}
        # And space-separated: {% img path width height %}
        image_path = parser.parse_expression()
        dimensions = []
        
        # Check if next token is comma (comma-separated syntax)
        if parser.stream.current.test('comma'):
            while parser.stream.skip_if('comma'):
                dimensions.append(parser.parse_expression())
        else:
            # Space-separated syntax: read remaining integers until block_end
            while not parser.stream.current.test('block_end'):
                dimensions.append(parser.parse_expression())
        
        args = [image_path, nodes.List(dimensions)]
        body = []
        return nodes.CallBlock(self.call_method('_image', args),
                               [], [], body).set_lineno(lineno)

    def _image(self, image_path, dimensions, caller):
        """Insert an image at the current cell position.
        
        Args:
            image_path: Path to the image file
            dimensions: List of optional [width, height] in pixels
        """
        if not pil:
            return ''
        
        # Extract width and height from dimensions list
        width = dimensions[0] if len(dimensions) > 0 else None
        height = dimensions[1] if len(dimensions) > 1 else None
        
        # Create ImageRef with path and optional dimensions
        image_ref = ImageRef(str(image_path), 0, width, height)
        if image_ref.image:
            node = self.environment.node_map.current_node
            if hasattr(node, 'current_cell'):
                # Attach to the current cell so coordinates are filled when the cell is written
                cell = node.current_cell
                cell.image_ref = image_ref
            else:
                # Fallback: keep existing behavior if no current cell is available
                node.set_image_ref(image_ref)
        return ''
