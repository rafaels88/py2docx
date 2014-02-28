# coding: utf-8
from elements.text import Paragraph
from elements.image import Image
from util import Unit


class CellAppendException(Exception):
    def __init__(self, value):
        self.value = value

    def __str__(self):
        return repr(self.value)


class AddColumnException(Exception):
    def __init__(self, value):
        self.value = value

    def __str__(self):
        return repr(self.value)


class InvalidCellMarginException(Exception):
    def __init__(self, value):
        self.value = value

    def __str__(self):
        return repr(self.value)


class Table(object):

    def __init__(self):
        self.xml_string = '\n<ns0:tbl>{content}</ns0:tbl>'
        self.columns = []

    def add_column(self, cells):
        xml_column = '\n<ns0:tr>{cells}</ns0:tr>'
        xml_cells = ''
        for cell in cells:
            if isinstance(cell, Cell):
                xml_cells += cell._get_xml()
            else:
                raise AddColumnException("Must be a list of Cell")
        self.columns.append(xml_column.format(cells=xml_cells))

    def _get_xml(self):
        return self.xml_string.format(content="".join(self.columns))


class Cell(object):

    def __init__(self, bgcolor=None, margin=None):
        self.content = ''
        self.bgcolor = bgcolor
        self.margin = margin
        self.xml_string = '\n<ns0:tc>' + \
                          '<ns0:tcPr>' + \
                          '{properties}' + \
                          '</ns0:tcPr>' + \
                          '{content}' + \
                          '</ns0:tc>'
        self.xml_props = []

        self._set_properties()

    def append(self, elem):
        if isinstance(elem, Paragraph) or isinstance(elem, Image):
            self.content += elem._get_xml()
        else:
            raise CellAppendException("Element to append should be a " +
                                      "Paragraph or Image")

    def _get_xml(self):
        return self.xml_string.replace('{content}', self.content)

    def _set_properties(self):
        self._set_bgcolor()
        self._set_margin()
        self.xml_string = self.xml_string.replace('{properties}', ''.join(self.xml_props))

    def _set_bgcolor(self):
        if self.bgcolor:
            xml = '<ns0:shd ns0:val="clear" ns0:color="auto" ns0:fill="{bgcolor}" />'

            self.xml_props.append(xml.format(bgcolor=self.bgcolor))

    def _set_margin(self):
        if self.margin:
            sizes = self.margin.split(' ')
            if len(sizes) == 1:
                top = bottom = left = right = sizes[0]
            elif len(sizes) == 2:
                top = bottom = sizes[0]
                left = right = sizes[1]
            elif len(sizes) == 3:
                top = sizes[0]
                right = left = sizes[1]
                bottom = sizes[2]
            elif len(sizes) == 4:
                top = sizes[0]
                right = sizes[1]
                bottom = sizes[2]
                left = sizes[3]
            else:
                raise InvalidCellMarginException("Cell Margin should be " +
                                                 "the W3C CSS format")

            top = int(top.rstrip('px'))
            bottom = int(bottom.rstrip('px'))
            right = int(right.rstrip('px'))
            left = int(left.rstrip('px'))

            xml = '<ns0:tcMar>' + \
                  '<ns0:start ns0:w="1440" ns0:type="dxa"/>' + \
                  '</ns0:tcMar>'

            self.xml_props.append(xml)
