# coding: utf-8
from py2docx.components.component import Component
from py2docx.templates.line_break_template import LineBreakTemplate


class LineBreak(Component):

    def __init__(self, *args, **kw):
        Component.__init__(self, *args, **kw)
        self.template = LineBreakTemplate()
