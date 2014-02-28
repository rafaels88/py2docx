# coding: utf-8
import os
import shutil
from datetime import datetime

TMP_PATH = '{0}/tmp'.format(os.path.dirname(os.path.abspath(__file__)))
DOCUMENT_PATH = '{0}/docx_document'.format(TMP_PATH)


class DocxDocument(object):

    def _create_file(self):
        dir_path = self._create_dir()
        obj_file = open("{0}{1}".format(dir_path, self.file_name), 'w')
        obj_file.write(self._get_xml())
        obj_file.close()

    def _create_dir(self):
        dir_to_create = "{0}/{1}/".format(DOCUMENT_PATH,
                                          self.dir_name.rstrip('/'))
        if not os.path.exists(dir_to_create):
            os.makedirs(dir_to_create)
        return dir_to_create

    @staticmethod
    def _clean_all():
        if os.path.exists(DOCUMENT_PATH):
            shutil.rmtree(DOCUMENT_PATH)
        os.makedirs(DOCUMENT_PATH)


class DocumentRelationshipFile(DocxDocument):

    def __init__(self):
        self.dir_name = 'word/_rels'
        self.file_name = 'document.xml.rels'
        self.xml_string = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + \
                          '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + \
                          '{content}' + \
                          '</Relationships>'
        self.rels = []

        self.IMAGE_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
        self.REL_PROTO = '<Relationship Id="{id}" Type="{type}" Target="{target}"/>'

    def _get_xml(self):
        rels_xml = "".join(self.rels)
        return self.xml_string.format(content=rels_xml)

    def _add_image(self, file_name):
        rel_id = "rId{0}".format(len(self.rels)+1)
        rel = self.REL_PROTO.format(id=rel_id,
                                    type=self.IMAGE_TYPE,
                                    target="media/{0}".format(file_name))
        self.rels.append(rel)
        return rel_id


class RelationshipFile(DocxDocument):

    def __init__(self):
        self.dir_name = '_rels'
        self.file_name = '.rels'
        self.xml_string = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + \
                          '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + \
                          '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>' + \
                          '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' + \
                          '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' + \
                          '</Relationships>'

    def _get_xml(self):
        return self.xml_string


class AppFile(DocxDocument):

    def __init__(self):
        self.dir_name = 'docProps'
        self.file_name = 'app.xml'
        self.xml_string = '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">' + \
                            '<Template>Normal.dotm</Template>'+ \
                            '<TotalTime>1</TotalTime>'+ \
                            '<Pages>1</Pages>'+ \
                            '<Words>3</Words>'+ \
                            '<Characters>23</Characters>'+ \
                            '<Application>Microsoft Office Word</Application>'+ \
                            '<DocSecurity>0</DocSecurity>'+ \
                            '<Lines>1</Lines>'+ \
                            '<Paragraphs>1</Paragraphs>'+ \
                            '<ScaleCrop>false</ScaleCrop>'+ \
                            '<Company></Company>'+ \
                            '<LinksUpToDate>false</LinksUpToDate>'+ \
                            '<CharactersWithSpaces>25</CharactersWithSpaces>'+ \
                            '<SharedDoc>false</SharedDoc>'+ \
                            '<HyperlinksChanged>false</HyperlinksChanged>'+ \
                            '<AppVersion>12.0000</AppVersion>'+ \
                        '</Properties>'

    def _get_xml(self):
        return self.xml_string


class CoreFile(DocxDocument):

    def __init__(self):
        self.dir_name = 'docProps'
        self.file_name = 'core.xml'
        self.props = {}
        self.xml_string = '<ns0:coreProperties xmlns:ns0="http://schemas.openxmlformats.org/package/2006/metadata/core-properties">' + \
                            '<dc:title xmlns:dc="http://purl.org/dc/elements/1.1/">{title}</dc:title>' + \
                            '<dc:subject xmlns:dc="http://purl.org/dc/elements/1.1/">{subject}</dc:subject>' + \
                            '<dc:creator xmlns:dc="http://purl.org/dc/elements/1.1/">{creator}</dc:creator>' + \
                            '<ns0:keywords>{keywords}</ns0:keywords>' + \
                            '<ns0:lastModifiedBy>{creator}</ns0:lastModifiedBy>' + \
                            '<ns0:revision>1</ns0:revision>' + \
                            '<ns0:category></ns0:category>' + \
                            '<dc:description xmlns:dc="http://purl.org/dc/elements/1.1/">{description}</dc:description>' + \
                            '<dcterms:created xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:dcterms="http://purl.org/dc/terms/" xsi:type="dcterms:W3CDTF">{date_created}</dcterms:created>' + \
                            '<dcterms:modified xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:dcterms="http://purl.org/dc/terms/" xsi:type="dcterms:W3CDTF">{date_created}</dcterms:modified>' + \
                          '</ns0:coreProperties>'

    def _get_xml(self):
        date_created = datetime.strftime(datetime.today(), '%Y-%m-%dT%H:%M:%SZ')
        return self.xml_string.format(title=self.props.get('title', ''),
                                      subject=self.props.get('subject', ''),
                                      creator=self.props.get('creator', ''),
                                      keywords=self.props.get('keywords', ''),
                                      description=self.props.get('description',
                                                                ''),
                                      date_created=date_created)


class DocumentFile(DocxDocument):

    def __init__(self):
        self.dir_name = 'word'
        self.file_name = 'document.xml'
        self.content = ''
        self.xml_string = '<ns0:document xmlns:ns0="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' + \
                            '<ns0:body>{content}</ns0:body>' + \
                          '</ns0:document>'

    def _get_xml(self):
        return self.xml_string.format(content=self.content)

    def _set_content(self, content):
        self.content = content

class ContentTypeFile(DocxDocument):

    def __init__(self):
        self.dir_name = ''
        self.file_name = '[Content_Types].xml'
        self.xml_string = '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' + \
                           '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' + \
                           '<Default Extension="xml" ContentType="application/xml"/>' + \
                           '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' + \
                           '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' + \
                           '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>' + \
                           '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>' + \
                           '<Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>' + \
                           '<Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>' + \
                           '<Override PartName="/word/webSettings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml"/>' + \
                           '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' + \
                           '</Types>'

    def _get_xml(self):
        return self.xml_string
