# coding: utf-8
from py2docx.components.container_component import ContainerComponent
from py2docx.templates.paragraph_template import ParagraphTemplate


class Paragraph(ContainerComponent):

    def __init__(self, *args, **kw):
        ContainerComponent.__init__(self, *args, **kw)
        self.template = ParagraphTemplate()

    def align(self, value):
        if value in ['left', 'right', 'center', 'justify']:
            self.template.align(value)
        return self
