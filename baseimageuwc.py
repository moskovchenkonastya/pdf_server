
# Needed on case-insensitive filesystems
from __future__ import absolute_import

# Try to import PIL in either of the two ways it can be installed.
try:
    from PIL import Image, ImageDraw
except ImportError:  # pragma: no cover
    import Image
    import ImageDraw

import qrcode.image.base


class UWCImage(qrcode.image.base.BaseImage):
    """
    PIL image builder, default format is PNG.
    """
    kind = "PNG"

    def new_image(self, **kwargs):
        mask = Image.open("C:\\enviroment\\qr\\mask.png")

        self.mask = mask.load()
        self.mask_size = mask.size

        back_color = kwargs.get("fill_color", "white")
        fill_color = kwargs.get("back_color", "black")

        mode = "RGB"

        img = Image.new(mode, (self.pixel_size, self.pixel_size), back_color)
        self.fill_color = fill_color
        self._idr = ImageDraw.Draw(img)
        return img

    def drawrect(self, row, col):

        t_width = self.mask_size[0] / float(self.width)
        t_height = self.mask_size[1] / float(self.width)
        t_mid = (self.mask_size[0] + self.mask_size[1]) / 2

        mask_x = min( int(t_width * col), self.mask_size[0]) 
        mask_y = int(t_width * row) 

        mask_y = mask_y - (t_mid - (self.mask_size[1]))

        box = self.pixel_box(row, col)


        if ( mask_y < 0 or mask_y >= self.mask_size[1] or self.mask[mask_x, mask_y] == (255, 255, 255)):
            self._idr.rectangle(box, fill=self.fill_color )
        else:
            #raw_color = min((self.mask[mask_x, mask_y][0] + self.mask[mask_x, mask_y][0] + self.mask[mask_x, mask_y][0]) / 3,150)

            #fill_color = (raw_color, raw_color, raw_color)
            fill_color=self.mask[mask_x, mask_y]


            self._idr.rectangle(box, fill=fill_color)

    def save(self, stream, format=None, **kwargs):
        if format is None:
            format = kwargs.get("kind", self.kind)
        if "kind" in kwargs:
            del kwargs["kind"]
        self._img.save(stream, format=format, **kwargs)

    def __getattr__(self, name):
        return getattr(self._img, name)
