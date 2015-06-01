# coding: utf-8
from py2docx.templates.component_template import ComponentTemplate


class BreakLineTemplate(ComponentTemplate):

    def __init__(self, *args, **kw):
        super(BreakLineTemplate, self).__init__(*args, **kw)

    def begin(self):
        return '<w:br/>'

    def properties(self):
        return ''

    def end(self):
        return ''
