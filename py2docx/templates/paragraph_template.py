# coding: utf-8
from py2docx.templates.component_template import ComponentTemplate


class ParagraphTemplate(ComponentTemplate):

    def __init__(self, *args, **kw):
        super(ParagraphTemplate, self).__init__(*args, **kw)
        self.properties_list = []
        self.align_property = ''

    def begin(self):
        return '<w:p>'

    def properties(self):
        self._build_properties_list()
        xml = '<w:pPr>'
        xml += ''.join(self.properties_list)
        xml += '</w:pPr>'
        return xml

    def end(self):
        return '</w:p>'

    def align(self, value):
        self.align_property = '<w:jc w:val="{0}"/>'.format(value)

    def _build_properties_list(self):
        self.properties_list.append(self.align_property)
