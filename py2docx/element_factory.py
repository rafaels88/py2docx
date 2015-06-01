# coding: utf-8
from py2docx.elements.paragraph import Paragraph


class ElementFactory(object):

    def new_paragraph(self):
        return Paragraph()
