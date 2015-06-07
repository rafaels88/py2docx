# coding: utf-8
from py2docx.components.component import Component
from py2docx.components.image import Image


class Document(Component):

    def __init__(self, *args, **kw):
        Component.__init__(self, *args, **kw)
        self.components = []
        self.images = []

    def add_component(self, component):
        self.components.append(component)
        if isinstance(component, Image):
            self.images.append(component)

    def retrieve_images(self):
        return self.images

    def draw(self):
        drew_components = map(lambda c: c.draw(), self.components)
        return ''.join(drew_components)
