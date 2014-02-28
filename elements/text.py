# coding: utf-8


class IsNotTextTypeException(Exception):
    def __init__(self, value):
        self.value = value

    def __str__(self):
        return repr(self.value)


class Text(object):
    pass


class Break(Text):

    def __init__(self):
        self.xml = """<ns0:br/>"""

    def _get_xml(self):
        return self.xml


class InlineText(Text):

    def __init__(self, text):
        self.text = text
        self.xml = """<ns0:r><ns0:t>{text}</ns0:t></ns0:r>"""

    def _get_xml(self):
        return self.xml.format(text=self.text)


class Paragraph(Text):

    def __init__(self):
        self.content = ""
        self.xml = """<ns0:p>{content}</ns0:p>"""

    def _get_xml(self):
        return self.xml.format(content=self.content)

    def append(self, elem):
        if isinstance(elem, Text):
            self.content += elem._get_xml()
        else:
            raise IsNotTextTypeException("Element to append should be a " +
                                         "Text Type")
