# coding: utf-8
from py2docx.structure_files import StructureFile


class App(StructureFile):

    def __init__(self):
        self.dir_name = 'docProps'
        self.file_name = 'app.xml'

    def draw(self):
        xml_string = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + \
                     '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" ' + \
                     'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">' + \
                     '<Template>Normal.dotm</Template>' + \
                     '<TotalTime>0</TotalTime>' + \
                     '<Pages>1</Pages>' + \
                     '<Words>0</Words>' + \
                     '<Characters>0</Characters>' + \
                     '<Application></Application>' + \
                     '<DocSecurity>0</DocSecurity>' + \
                     '<Lines>1</Lines>' + \
                     '<Paragraphs>1</Paragraphs>' + \
                     '<ScaleCrop>false</ScaleCrop>' + \
                     '<LinksUpToDate>false</LinksUpToDate>' + \
                     '<CharactersWithSpaces>0</CharactersWithSpaces>' + \
                     '<SharedDoc>false</SharedDoc>' + \
                     '<HyperlinksChanged>false</HyperlinksChanged>' + \
                     '<AppVersion>12.0000</AppVersion>' + \
                     '</Properties>'
        return xml_string
