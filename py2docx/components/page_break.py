# coding: utf-8
from py2docx.components.component import Component
from py2docx.templates.page_break_template import PageBreakTemplate


class PageBreak(Component):

    def __init__(self, *args, **kw):
        Component.__init__(self, *args, **kw)
        self.template = PageBreakTemplate()
