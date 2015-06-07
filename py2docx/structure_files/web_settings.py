# coding: utf-8
from py2docx.structure_files import StructureFile


class WebSettings(StructureFile):

    def __init__(self):
        self.dir_name = 'word'
        self.file_name = 'webSettings.xml'

    def draw(self):
        xml_string = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + \
                     '<w:webSettings ' + \
                     'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' + \
                     'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' + \
                     '<w:allowPNG/>' + \
                     '<w:doNotSaveAsSingleFile/>' + \
                     '</w:webSettings>'
        return xml_string
