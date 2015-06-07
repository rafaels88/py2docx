# coding: utf-8
from py2docx.templates.component_template import ComponentTemplate


class PageBreakTemplate(ComponentTemplate):

    def __init__(self, *args, **kw):
        super(PageBreakTemplate, self).__init__(*args, **kw)

    def begin(self):
        return '<w:br w:type="page"/>'

    def properties(self):
        return ''

    def end(self):
        return ''
