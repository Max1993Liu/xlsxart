from PIL import Image
import xlsxwriter
from collections import namedtuple


class RGB:

    def __init__(self, r, g, b):
        self.r = r
        self.g = g
        self.b = b

    @classmethod
    def BLACK(cls):
        return cls(0, 0, 0)

    @classmethod
    def WHITE(cls):
        return cls(255, 255, 255)

    def to_hex(self):
        """ Convert to hex color"""
        return '#'+''.join([("%0.2X" % i).upper() for i in [self.r, self.g, self.b]])


# note that in PIL x and y starts from the upper left corner
Pixel = namedtuple('Pixel', ['x', 'y', 'color'])


class Picture:
    """ A picture is basically a bunch of pixels with color"""

    def __init__(self, path):
        self.img = Image.open(path)
        self.img.thumbnail((1024, 800), Image.ANTIALIAS)

    def size(self):
        return self.img.size

    def to_pixels(self):
        width, height = self.img.size
        all_pixels = self.img.load()
        return [Pixel(x, y, RGB(*all_pixels[x, y]))
                for x in range(width) for y in range(height)]


class XLSXCanvas:

    def __init__(self, path):
        self.workbook = xlsxwriter.Workbook(path)

    def draw(self, picture: Picture):
        pic_size = picture.size()
        worksheet = self.create_new_sheet('paint', pic_size)
        for pixel in picture.to_pixels():
            self.draw_pixel(worksheet, pixel)
        self.finish()

    def create_new_sheet(self, name, size=(1000, 1000)):
        worksheet = self.workbook.add_worksheet(name)
        # make each cell into little square
        worksheet.set_column(0, size[0], 1)
        worksheet.set_default_row(7)
        return worksheet

    def draw_pixel(self, worksheet, pixel: Pixel):
        worksheet.write(pixel.y, pixel.x, None, self.create_cell_format(pixel.color))

    def create_cell_format(self, rgb_color: RGB):
        """ Given an RGB color instance create a new cell formatter
            with the corresponding background color
        """
        return self.workbook.add_format({'bg_color': rgb_color.to_hex()})

    def finish(self):
        self.workbook.close()


canvas = XLSXCanvas('./1.xlsx')
pic = Picture('./image/1.JPG')
canvas.draw(pic)






