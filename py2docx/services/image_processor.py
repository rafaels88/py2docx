# coding: utf-8
import os
from os.path import basename
from PIL import Image as PILImage
from py2docx.config import DOCUMENT_PATH
from py2docx.helpers.unit_conversor import UnitConversor


class ImageProcessor(object):

    def __init__(self, image_path):
        self.file = open(image_path, 'rb')
        self._upload_image()
        self.image = PILImage.open(self.file.name)

    def height_by_percentage(self, percentage):
        pixels = self._height_in_pixels()
        return self._real_size_by_pixels_and_percentage(pixels, percentage)

    def width_by_percentage(self, percentage):
        pixels = self._width_in_pixels()
        return self._real_size_by_pixels_and_percentage(pixels, percentage)

    def image_name(self):
        return basename(self.file.name)

    def image_path(self):
        return "media/{0}".format(self.image_name())

    def _real_size_by_pixels_and_percentage(self, pixels, percentage):
        size = UnitConversor.pixel_to_emu(pixels)
        percentage = float(percentage.rstrip('%'))
        real_size = (percentage / 100) * size
        return int(real_size)

    def _width_in_pixels(self):
        return self.image.size[1]

    def _height_in_pixels(self):
        return self.image.size[0]

    def _upload_image(self):
        dir_media = "{0}/word/media".format(DOCUMENT_PATH)
        if not os.path.exists(dir_media):
            os.makedirs(dir_media)
        uploaded = open("{0}/{1}".format(dir_media, self.image_name()), 'wb')
        uploaded.write(self.file.read())
        uploaded.close()
