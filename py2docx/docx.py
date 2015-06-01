# coding: utf-8

from py2docx.components.docx_component import DocxComponent


class Docx(object):

    def __init__(self):
        self.docx_component = DocxComponent()

    def append(self, component):
        self.docx_component.add_component(component)
        return self

    def save(self):
        return self.docx_component.draw()
