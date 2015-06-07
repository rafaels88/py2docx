# coding: utf-8
from py2docx.structure_files import StructureFile


class Document(StructureFile):

    def __init__(self):
        self.dir_name = 'word'
        self.file_name = 'document.xml'
        self.content = ''

    def draw(self):
        xml_string = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + \
                     '<w:document ' + \
                     'xmlns:ve="http://schemas.openxmlformats.org/markup-compatibility/2006" ' + \
                     'xmlns:o="urn:schemas-microsoft-com:office:office" ' + \
                     'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' + \
                     'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" ' + \
                     'xmlns:v="urn:schemas-microsoft-com:vml" ' + \
                     'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' + \
                     'xmlns:w10="urn:schemas-microsoft-com:office:word" ' + \
                     'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' + \
                     'xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml">' + \
                     '  <w:body>{0}</w:body>' + \
                     '</w:document>'
        return xml_string.format(self.content)

    def set_document_component(self, component):
        self.content = component.draw()
