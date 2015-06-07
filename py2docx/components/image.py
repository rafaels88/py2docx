# coding: utf-8
from py2docx.components.component import Component
from py2docx.templates.image_template import ImageTemplate
from py2docx.services.image_processor import ImageProcessor


class Image(Component):

    def __init__(self, image_path, *args, **kw):
        Component.__init__(self, *args, **kw)
        self.image_processor = ImageProcessor(image_path)
        self.template = ImageTemplate(self.image_processor.image_name())
        self.width('100%')
        self.height('100%')

    def set_relationship_id(self, id):
        self.template.set_relationship_id(id)

    def image_name(self):
        return self.image_processor.image_name()

    def image_path(self):
        return self.image_processor.image_path()

    def align(self, value):
        if value in ['left', 'right', 'center', 'justify']:
            self.template.align(value)
        return self

    def width(self, percentage):
        width = self.image_processor.width_by_percentage(percentage)
        self.template.width(width)
        return self

    def height(self, percentage):
        height = self.image_processor.height_by_percentage(percentage)
        self.template.height(height)
        return self
