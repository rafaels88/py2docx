# coding: utf-8
import re
from py2docx.components.component import Component
from py2docx.templates.text_template import TextTemplate


class Text(Component):

    def __init__(self, text, *args, **kw):
        Component.__init__(self, *args, **kw)
        self.template = TextTemplate()
        self.template.append(text)

    def bold(self):
        self.template.bold()
        return self

    def italic(self):
        self.template.italic()
        return self

    def underline(self, style='solid', color='000000'):
        if style in ['dotted', 'dashed', 'solid', 'double']:
            color = color.lstrip('#')
            if style == 'dashed':
                style = 'dash'
            elif style == 'solid':
                style = 'single'
            self.template.underline(style, color)
        return self

    def uppercase(self):
        self.template.uppercase()
        return self

    def color(self, color):
        self.template.color(color)
        return self

    def font(self, font):
        if font in ['Cambria', 'Times New Roman',
                    'Arial', 'Calibri']:
            self.template.font(font)
        return self

    def size(self, size):
        if type(size) in [str, unicode]:
            size = re.sub("\D", "", size)
            if size:
                size = int(size)

        if type(size) is int:
            size *= 2
            self.template.size(size)
        return self

    def block(self):
        self.template.block()
        return self
