# coding: utf-8
from datetime import datetime
from py2docx.structure_files import StructureFile


class Core(StructureFile):

    def __init__(self):
        self.dir_name = 'docProps'
        self.file_name = 'core.xml'
        self.title = ''
        self.creator = ''
        self.keywords = ''
        self.subject = ''

    def title(self, value):
        self.title = value

    def creator(self, value):
        self.creator = value

    def keywords(self, value):
        self.keywords = value

    def subject(self, value):
        self.subject = value

    def draw(self):
        xml_string = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + \
                      '<cp:coreProperties ' + \
                      'xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ' + \
                      'xmlns:dc="http://purl.org/dc/elements/1.1/" ' + \
                      'xmlns:dcterms="http://purl.org/dc/terms/" ' + \
                      'xmlns:dcmitype="http://purl.org/dc/dcmitype/" ' + \
                      'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' + \
                      '<dc:title>{title}</dc:title>' + \
                      '<dc:subject>{subject}</dc:subject>' + \
                      '<dc:creator>{creator}</dc:creator>' + \
                      '<cp:keywords>{keywords}</cp:keywords>' + \
                      '<cp:lastModifiedBy>{creator}</cp:lastModifiedBy>' + \
                      '<cp:revision>1</cp:revision>' + \
                      '<dcterms:created xsi:type="dcterms:W3CDTF">{created_at}</dcterms:created>' + \
                      '<dcterms:modified xsi:type="dcterms:W3CDTF">{created_at}</dcterms:modified>' + \
                      '</cp:coreProperties>'

        return xml_string.format(title=self.title,
                                 subject=self.subject,
                                 creator=self.creator,
                                 keywords=self.keywords,
                                 created_at=self._created_at())

    def _created_at(self):
        return datetime.strftime(datetime.today(),
                                 '%Y-%m-%dT%H:%M:%SZ')
