# coding: utf-8
from py2docx.elements import Block


class IsNotTextTypeException(Exception):
    def __init__(self, value):
        self.value = value

    def __str__(self):
        return repr(self.value)


class Text(object):
    pass


class Break(Text):

    def __init__(self):
        self.xml = """<w:br/>"""

    def _get_xml(self):
        return self.xml


class InlineText(Text):

    def __init__(self, text, bold=False, italic=False, underline=None,
                 uppercase=False, color=None):
        self.text = text
        self.bold = bold
        self.italic = italic
        self.underline = underline
        self.uppercase = uppercase
        self.color = color
        self.xml_string = '<w:r>' + \
                          '<w:rPr>' + \
                          '{properties}' + \
                          '</w:rPr>' + \
                          '<w:t xml:space="preserve">' + \
                          '{text}' + \
                          '</w:t>' + \
                          '</w:r>'
        self.xml_props = []
        self._set_properties()

    def _get_xml(self):
        if type(self.text) not in [str, unicode]:
            self.text = unicode(self.text)
        return self.xml_string.replace("{text}",
                                       self.text.encode('utf-8'))

    def _set_properties(self):
        self._set_bold()
        self._set_italic()
        self._set_underline()
        self._set_uppercase()
        self._set_color()
        self.xml_string = self.xml_string.replace('{properties}',
                                                  ''.join(self.xml_props))

    def _set_bold(self):
        if self.bold is True:
            xml = '<w:b w:val="true"/>'
            self.xml_props.append(xml)

    def _set_italic(self):
        if self.italic is True:
            xml = '<w:i w:val="true"/>'
            self.xml_props.append(xml)

    def _set_underline(self):
        if type(self.underline) is dict:
            if self.underline.get('style', None) in ['dotted', 'dashed',
                                                     'solid', 'double']:
                style = self.underline.get('style')
                color = self.underline.get('color', '000000').lstrip('#')
                if style == 'dashed':
                    style = 'dash'
                elif style == 'solid':
                    style = 'single'

            xml = '<w:u w:val="{type}" w:color="{color}"/>'.format(type=style,
                                                                   color=color)
            self.xml_props.append(xml)

    def _set_uppercase(self):
        if self.uppercase is True:
            xml = '<w:caps w:val="true" />'
            self.xml_props.append(xml)

    def _set_color(self):
        if type(self.color) in [str, unicode]:
            xml = '<w:color w:val="{color}" />'
            color = self.color.lstrip('#')
            self.xml_props.append(xml.format(color=color))


class BlockText(InlineText):

    def _get_xml(self):
        block_text = Block()
        block_text.append(InlineText(self.text))
        return block_text._get_xml()
