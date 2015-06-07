# coding: utf-8
from py2docx.components.component import Component
from py2docx.templates.break_line_template import BreakLineTemplate


class BreakLine(Component):

    def __init__(self, *args, **kw):
        Component.__init__(self, *args, **kw)
        self.template = BreakLineTemplate()
