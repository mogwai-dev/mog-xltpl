from jinja2 import pass_environment
import hashlib
import os
from datetime import datetime

def add_filter(filter):
    def wrapper(env, *args):
        cell_node = env.node_map.current_node.current_cell
        cell_node.add_filter(filter, args)
        return ''
    return wrapper


class ImageMarker:
    """Marker object to represent an image in template rendering.
    
    This allows {{ }} syntax to work with images by returning a special
    object that gets processed in the finalize hook.
    """
    def __init__(self, path, width=None, height=None):
        self.path = path
        self.width = width
        self.height = height
        self._is_image_marker = True
    
    def __str__(self):
        # Return empty string so cell appears empty in Excel
        return ''
    
    def __repr__(self):
        return f'ImageMarker({self.path}, {self.width}x{self.height})'


@pass_environment
def img_filter(env, image_path, width=None, height=None):
    """Image filter for embedding images in Excel cells.
    
    Supports both {{ }} and {% xv %} syntax:
        {{ image_path | img(150, 120) }}        # Recommended - standard Jinja2
        {% xv image_path | img(150, 120) %}     # Legacy - still works
    
    Returns ImageMarker for {{ }} syntax, or stores in env for {% xv %} syntax.
    """
    import os
    import six
    
    if not image_path:
        return ''
    
    # Check if file exists
    try:
        from PIL.ImageFile import ImageFile
        if isinstance(image_path, ImageFile):
            path = image_path
        else:
            fname = six.text_type(image_path)
            if not os.path.exists(fname):
                return ''
            path = fname
    except ImportError:
        return ''
    
    # Return ImageMarker for {{ }} syntax
    # The finalize hook will process this
    return ImageMarker(path, width, height)


def sha256_filter(file_path):
    """Calculate SHA256 hash of a file.
    
    Usage in template:
        {{ file_path | sha256 }}
    
    Returns the hex digest of the file's SHA256 hash.
    Raises an error if file doesn't exist or cannot be read.
    """
    if not file_path:
        raise ValueError("sha256 filter requires a file path")
    
    file_path = str(file_path)
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"File not found: {file_path}")
    
    try:
        sha256_hash = hashlib.sha256()
        with open(file_path, 'rb') as f:
            for byte_block in iter(lambda: f.read(4096), b''):
                sha256_hash.update(byte_block)
        return sha256_hash.hexdigest()
    except FileNotFoundError:
        raise
    except Exception as e:
        raise RuntimeError(f"Error reading file {file_path}: {e}")


def mtime_filter(file_path, format_str=None):
    """Get file modification time.
    
    Usage in template:
        {{ file_path | mtime }}                              # Default format: ISO 8601
        {{ file_path | mtime('%Y-%m-%d %H:%M:%S') }}        # Custom format
        {{ file_path | mtime('%Y%m%d') }}                   # YYYYMMDD format
    
    Returns formatted modification time.
    Raises an error if file doesn't exist or cannot be accessed.
    Default format is ISO 8601: YYYY-MM-DDTHH:MM:SS
    """
    if not file_path:
        raise ValueError("mtime filter requires a file path")
    
    file_path = str(file_path)
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"File not found: {file_path}")
    
    try:
        mtime_timestamp = os.path.getmtime(file_path)
        mtime_datetime = datetime.fromtimestamp(mtime_timestamp)
        
        if format_str:
            return mtime_datetime.strftime(format_str)
        else:
            # Default ISO 8601 format
            return mtime_datetime.isoformat()
    except FileNotFoundError:
        raise
    except Exception as e:
        raise RuntimeError(f"Error getting mtime for {file_path}: {e}")


def to_fullwidth_filter(value):
    """Convert half-width digits to full-width and '-' to '－'.

    Usage in template:
        {{ value | to_fullwidth }}

    Converts:
        - 0-9 to ０-９
        - '-' (U+002D) to '－' (U+FF0D, fullwidth hyphen-minus)

    Args:
        value: String or number to convert

    Returns converted string, or empty string if value is not a string.
    """
    if not isinstance(value, str):
        if value is None:
            return ''
        value = str(value)

    # Translation table: half-width to full-width
    translation_map = str.maketrans(
        '0123456789-',
        '０１２３４５６７８９－'
    )
    return value.translate(translation_map)
