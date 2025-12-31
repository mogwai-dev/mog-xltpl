from copy import deepcopy
from io import BytesIO
from openpyxl.drawing.image import Image

# to avoid file closed error
class Cache():

    def __init__(self):
        self.map = {}

    def get_data(self, key):
        return self.map.get(key)

    def set_data(self, key, data):
        self.map[key] = data

    def clear(self):
        self.map.clear()


img_cache = Cache()
data_cache = Cache()

class Img(Image):
    """Wrapper around openpyxl Image to properly handle file references"""

    def __init__(self, image):
        self._cached_data = None
        # If image is already an Image object, rebuild it from in-memory data to avoid broken paths
        if isinstance(image, Image):
            try:
                data = image._data()
                self._cached_data = data
                super(Img, self).__init__(BytesIO(data))
            except Exception:
                # If image.path is a relative Excel path (e.g., /xl/media/image1.jpeg), skip it
                # These are internal Excel references that can't be loaded directly
                if hasattr(image, 'path') and isinstance(image.path, str):
                    path = image.path
                    # Check if it's a valid file path, not an Excel internal reference
                    import os
                    if not path.startswith('/') and os.path.exists(path):
                        super(Img, self).__init__(path)
                    else:
                        # Fallback: create empty image data to avoid crash
                        raise ValueError(f"Cannot load image from internal Excel path: {path}")
                else:
                    raise
            # Copy relevant attributes
            self.anchor = deepcopy(image.anchor) if hasattr(image, 'anchor') else None
            self.width = image.width if hasattr(image, 'width') else None
            self.height = image.height if hasattr(image, 'height') else None
        else:
            # Assume it's a file path
            with open(image, 'rb') as f:
                data = f.read()
            self._cached_data = data
            super(Img, self).__init__(BytesIO(data))

    def _data(self):
        if self._cached_data is not None:
            return self._cached_data
        return super(Img, self)._data()



