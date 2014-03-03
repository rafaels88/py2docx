# coding: utf-8
from elements import Block
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

    def __init__(self, margin=None):
        self.xml_string = '<w:tbl>' + \
                          '<w:tblPr>' + \
                          '{properties}' + \
                          '</w:tblPr>' + \
                          '{content}' + \
                          '</w:tbl>'
        self.margin = margin
        self.columns = []
        self.xml_props = []
        self._set_properties()

    def add_column(self, cells):
        xml_column = '\n<w:tr>{cells}</w:tr>\n'
        xml_cells = ''
        for cell in cells:
            if isinstance(cell, Cell):
                xml_cells += cell._get_xml()
            else:
                raise AddColumnException("Must be a list of Cell")
        self.columns.append(xml_column.format(cells=xml_cells))

    def _set_properties(self):
        self._set_margin()
        self.xml_string = self.xml_string.replace('{properties}',
                                                  ''.join(self.xml_props))

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
                                                 "in the W3C CSS format")

            top = Unit.to_dxa(top)
            bottom = Unit.to_dxa(bottom)
            right = Unit.to_dxa(right)
            left = Unit.to_dxa(left)

            xml = '\n<w:tblCellMar>' + \
                  '<w:top w:w="{top}" w:type="dxa" />' + \
                  '<w:left w:w="{left}" w:type="dxa" />' + \
                  '<w:bottom w:w="{bottom}" w:type="dxa" />' + \
                  '<w:right w:w="{right}" w:type="dxa" />' + \
                  '</w:tblCellMar>\n'

            self.xml_props.append(xml.format(top=top,
                                             bottom=bottom,
                                             right=right,
                                             left=left))

    def _get_xml(self):
        return self.xml_string.format(content="".join(self.columns))


class Cell(object):

    def __init__(self, bgcolor=None, margin=None, width=None,
                 valign=None, nowrap=False):
        self.content = ''
        self.bgcolor = bgcolor
        self.margin = margin
        self.width = width
        self.valign = valign
        self.nowrap = nowrap
        self.xml_string = '\n<w:tc>' + \
                          '<w:tcPr>' + \
                          '{properties}' + \
                          '</w:tcPr>' + \
                          '{content}' + \
                          '</w:tc>'
        self.xml_props = []

        self._set_properties()

    def append(self, elem):
        if isinstance(elem, Block) \
           or isinstance(elem, Image) \
           or isinstance(elem, Table):
            self.content += elem._get_xml()
        else:
            raise CellAppendException("Element to append should be a " +
                                      "Block, Table or Image")

    def _get_xml(self):
        return self.xml_string.replace('{content}', self.content)

    def _set_properties(self):
        self._set_bgcolor()
        self._set_margin()
        self._set_width()
        self._set_valign()
        self._set_nowrap()
        self.xml_string = self.xml_string.replace('{properties}',
                                                  ''.join(self.xml_props))

    def _set_bgcolor(self):
        if self.bgcolor:
            xml = '\n<w:shd w:val="clear" w:color="auto" w:fill="{bgcolor}" />\n'
            self.xml_props.append(xml.format(bgcolor=self.bgcolor.lstrip('#')))

    def _set_width(self):
        if self.width:
            xml = '\n<w:tcW w:type="{unit_type}" w:w="{width}"/>\n'
            if self.width[-1:] == '%':
                unit_type = 'pct'
                xml = xml.format(unit_type=unit_type,
                                 width=self.width)

            elif self.width[-2:] in ['cm', 'pt', 'in']:
                unit_type = 'dxa'
                width = Unit.to_dxa(self.width)
                xml = xml.format(unit_type=unit_type,
                                 width=width)

            self.xml_props.append(xml)

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
                                                 "in the W3C CSS format")

            top = Unit.to_dxa(top)
            bottom = Unit.to_dxa(bottom)
            right = Unit.to_dxa(right)
            left = Unit.to_dxa(left)

            xml = '\n<w:tcMar>' + \
                  '<w:top w:w="{top}" w:type="dxa" />' + \
                  '<w:left w:w="{left}" w:type="dxa" />' + \
                  '<w:bottom w:w="{bottom}" w:type="dxa" />' + \
                  '<w:right w:w="{right}" w:type="dxa" />' + \
                  '</w:tcMar>\n'

            self.xml_props.append(xml.format(top=top,
                                             bottom=bottom,
                                             right=right,
                                             left=left))

    def _set_valign(self):
        if self.valign and type(self.valign) is str:
            self.valign = self.valign.lower()
            if self.valign == 'top' or \
               self.valign == 'center' or \
               self.valign == 'bottom':
                    xml = '\n<w:vAlign w:val="{align}" />\n'
                    self.xml_props.append(xml.format(align=self.valign))

    def _set_nowrap(self):
        if self.nowrap is True:
            xml = '<w:noWrap/>'
            self.xml_props.append(xml)
