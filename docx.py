# coding: utf-8
import os
import zipfile
from datetime import datetime


class Docx(object):

    def __init__(self):
        self.tmp_path = '{0}/tmp' \
                        .format(os.path.dirname(os.path.abspath(__file__)))
        self.app_path = '{0}/docx_document/docProps/app.xml' \
                        .format(self.tmp_path)
        self.core_path = '{0}/docx_document/docProps/core.xml' \
                         .format(self.tmp_path)
        self.document_path = '{0}/docx_document/word/document.xml' \
                             .format(self.tmp_path)
        self.content_type_path = '{0}/docx_document/[Content_Types].xml' \
                                 .format(self.tmp_path)
        self.rels_path = '{0}/docx_document/_rels/.rels'.format(self.tmp_path)
        self.props = {}
        self.body_content = ''

    def save(self):
        self._create_structure()
        zip_name = zipfile.ZipFile("{0}/file.docx".format(self.tmp_path), 'w')

        for dirpath, dirs, files in os.walk("{0}/docx_document"
                                            .format(self.tmp_path)):
            for f in files:
                file_name = os.path.join(dirpath, f)
                file_name_zip = file_name.replace('{0}/docx_document'
                                                  .format(self.tmp_path), '')
                zip_name.write(file_name, file_name_zip)

    def set_properties(self, props):
        self.props = props

    def append(self, elem):
        xml_elem = elem._get_xml()
        self.body_content += xml_elem

    def _create_structure(self):
        dirs_to_create = ['docx_document',
                          'docx_document/docProps',
                          'docx_document/word',
                          'docx_document/_rels']

        for dir_name in dirs_to_create:
            dir_to_create = "{0}/{1}".format(self.tmp_path, dir_name)
            if not os.path.exists(dir_to_create):
                os.makedirs(dir_to_create)

        self._create_app_file()
        self._create_content_type_file()
        self._create_core_file()
        self._create_document_file()
        self._create_rels_file()

    def _create_app_file(self):
        app_file = open(self.app_path, 'w')
        xml_string = """<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
    <Template>Normal.dotm</Template>
    <TotalTime>1</TotalTime>
    <Pages>1</Pages>
    <Words>3</Words>
    <Characters>23</Characters>
    <Application>Microsoft Office Word</Application>
    <DocSecurity>0</DocSecurity>
    <Lines>1</Lines>
    <Paragraphs>1</Paragraphs>
    <ScaleCrop>false</ScaleCrop>
    <Company></Company>
    <LinksUpToDate>false</LinksUpToDate>
    <CharactersWithSpaces>25</CharactersWithSpaces>
    <SharedDoc>false</SharedDoc>
    <HyperlinksChanged>false</HyperlinksChanged>
    <AppVersion>12.0000</AppVersion>
</Properties>"""

        app_file.write(xml_string)
        app_file.close()

    def _create_core_file(self):
        core_file = open(self.core_path, 'w')
        xml_string = """<ns0:coreProperties xmlns:ns0="http://schemas.openxmlformats.org/package/2006/metadata/core-properties">
  <dc:title xmlns:dc="http://purl.org/dc/elements/1.1/">{title}</dc:title>
  <dc:subject xmlns:dc="http://purl.org/dc/elements/1.1/">{subject}</dc:subject>
  <dc:creator xmlns:dc="http://purl.org/dc/elements/1.1/">{creator}</dc:creator>
  <ns0:keywords>{keywords}</ns0:keywords>
  <ns0:lastModifiedBy>{creator}</ns0:lastModifiedBy>
  <ns0:revision>1</ns0:revision>
  <ns0:category></ns0:category>
  <dc:description xmlns:dc="http://purl.org/dc/elements/1.1/">{description}</dc:description>
  <dcterms:created xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:dcterms="http://purl.org/dc/terms/" xsi:type="dcterms:W3CDTF">{date_created}</dcterms:created>
  <dcterms:modified xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:dcterms="http://purl.org/dc/terms/" xsi:type="dcterms:W3CDTF">{date_created}</dcterms:modified>
</ns0:coreProperties>"""

        date_created = datetime.strftime(datetime.today(), '%Y-%m-%dT%H:%M:%SZ')
        xml_string = xml_string.format(title=self.props.get('title', ''),
                                       subject=self.props.get('subject', ''),
                                       creator=self.props.get('creator', ''),
                                       keywords=self.props.get('keywords', ''),
                                       description=self.props.get('description',
                                                                  ''),
                                       date_created=date_created)

        core_file.write(xml_string)
        core_file.close()

    def _create_document_file(self):
        document_file = open(self.document_path, 'w')
        xml_string = """<ns0:document xmlns:ns0="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
   <ns0:body>{content}</ns0:body>
</ns0:document>"""
        document_file.write(xml_string.format(content=self.body_content))
        document_file.close()

    def _create_rels_file(self):
        rels_file = open(self.rels_path, 'w')
        xml_string = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""
        rels_file.write(xml_string)
        rels_file.close()

    def _create_content_type_file(self):
        ct_file = open(self.content_type_path, 'w')
        xml_string = """<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
   <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
   <Default Extension="xml" ContentType="application/xml"/>
   <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
   <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
   <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
   <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
   <Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
   <Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>
   <Override PartName="/word/webSettings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml"/>
   <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
</Types>"""
        ct_file.write(xml_string)
        ct_file.close()
