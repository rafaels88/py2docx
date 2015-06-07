# coding: utf-8
from py2docx.templates.component_template import ComponentTemplate


class TextTemplate(ComponentTemplate):

    def __init__(self, *args, **kw):
        super(TextTemplate, self).__init__(*args, **kw)
        self.properties_list = []
        self.bold_property = ''
        self.italic_property = ''
        self.underline_property = ''
        self.uppercase_property = ''
        self.color_property = ''
        self.font_property = ''
        self.size_property = ''
        self.is_block = False

    def begin(self):
        if self.is_block is True:
            return '<w:p><w:r>'
        else:
            return '<w:r>'

    def properties(self):
        self._build_properties_list()
        xml = '<w:pPr>'
        xml += ''.join(self.properties_list)
        xml += '</w:pPr>'
        return xml

    def contents(self):
        result = '<w:t xml:space="preserve">'
        result += ''.join(self.content_list)
        result += '</w:t>'
        return result

    def end(self):
        if self.is_block is True:
            return '</w:r></w:p>'
        else:
            return '</w:r>'

    def bold(self):
        self.bold_property = '<w:b w:val="true"/>'

    def italic(self):
        self.italic_property = '<w:i w:val="true"/>'

    def underline(self, style, color):
        self.underline_property = '<w:u w:val="{0}" w:color="{1}"/>' \
            .format(style, color)

    def uppercase(self):
        self.uppercase_property = '<w:caps w:val="true" />'

    def color(self, color):
        self.color_property = '<w:color w:val="{0}" />' \
            .format(color.lstrip('#'))

    def font(self, font):
        self.color_property = '<w:rFonts w:ascii="{font}" w:hAnsi="{font}"/>' \
            .format(font=font)

    def size(self, size):
        self.size_property = '<w:sz w:val="{0}"/>'.format(size)

    def block(self):
        self.is_block = True

    def _build_properties_list(self):
        properties = [
            self.bold_property,
            self.italic_property,
            self.underline_property,
            self.uppercase_property,
            self.color_property,
            self.font_property,
            self.size_property
        ]
        for property in properties:
            self.properties_list.append(property)
