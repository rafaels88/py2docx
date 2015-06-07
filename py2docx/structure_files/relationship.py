# coding: utf-8
from py2docx.structure_files import StructureFile


class Relationship(StructureFile):

    def __init__(self):
        self.dir_name = '_rels'
        self.file_name = '.rels'

    def draw(self):
        xml_string = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + \
                     '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + \
                     '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>' + \
                     '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' + \
                     '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' + \
                     '</Relationships>'
        return xml_string
