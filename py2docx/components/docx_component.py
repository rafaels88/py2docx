# coding: utf-8
from py2docx.components.component import Component


class DocxComponent(Component):

    def __init__(self, *args, **kw):
        Component.__init__(self, *args, **kw)
        self.components = []

    def add_component(self, component):
        self.components.append(component)

    def draw(self):
        drew_components = map(lambda c: c.draw(), self.components)
        return ''.join(drew_components)
