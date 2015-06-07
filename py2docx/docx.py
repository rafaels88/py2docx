# coding: utf-8

from py2docx.components.document import Document
from py2docx.services.docx_file_creator import DocxFileCreator


class Docx(object):

    def __init__(self):
        self.document = Document()

    def append(self, component):
        self.document.add_component(component)
        return self

    def save(self, path):
        docx_file_creator = DocxFileCreator(document=self.document)
        docx_file_creator.create(path)
