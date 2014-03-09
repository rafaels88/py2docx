# coding: utf-8
import os
import shutil
from datetime import datetime

TMP_PATH = '{0}/tmp'.format(os.path.dirname(os.path.abspath(__file__)))
DOCUMENT_PATH = '{0}'.format(TMP_PATH)


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
        obj_file = open("{0}/__init__.py".format(DOCUMENT_PATH), 'w')
        obj_file.close()


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
        self.HYPERLINK_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'
        self.REL_PROTO = '<Relationship Id="{id}" Type="{type}" Target="{target}"/>'
        self.REL_HYPERLINK = '<Relationship Id="{id}" Type="{type}" Target="{target}" TargetMode="External"/>'

    def _get_xml(self):
        rels_xml = "".join(self.rels)
        return self.xml_string.format(content=rels_xml)

    def _generate_rel_id(self):
        rel_id = "rId{0}".format(len(self.rels)+1)
        return rel_id

    def _add_image(self, file_name):
        rel_id = self._generate_rel_id()
        rel = self.REL_PROTO.format(id=rel_id,
                                    type=self.IMAGE_TYPE,
                                    target="media/{0}".format(file_name))
        self.rels.append(rel)
        return rel_id

    def _add_hyperlink(self, link):
        rel_id = self._generate_rel_id()
        rel = self.REL_HYPERLINK.format(id=rel_id,
                                        type=self.HYPERLINK_TYPE,
                                        target=link)
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
        self.xml_string = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + \
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

    def _get_xml(self):
        return self.xml_string


class CoreFile(DocxDocument):

    def __init__(self):
        self.dir_name = 'docProps'
        self.file_name = 'core.xml'
        self.props = {}
        self.xml_string = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + \
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
                          '<dcterms:created xsi:type="dcterms:W3CDTF">{date_created}</dcterms:created>' + \
                          '<dcterms:modified xsi:type="dcterms:W3CDTF">{date_created}</dcterms:modified>' + \
                          '</cp:coreProperties>'

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
        self.xml_string = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + \
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
                          '  <w:body>{content}</w:body>' + \
                          '</w:document>'

    def _get_xml(self):
        return self.xml_string.format(content=self.content)

    def _set_content(self, content):
        self.content = content

class ContentTypeFile(DocxDocument):

    def __init__(self):
        self.dir_name = ''
        self.file_name = '[Content_Types].xml'
        self.xml_string = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + \
                          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' + \
                          '<Default Extension="xml" ContentType="application/xml"/>' + \
                          '<Override PartName="/word/fontTable.xml" ' + \
                          'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>' + \
                          '<Default ContentType="image/jpeg" Extension="jpg"/>' + \
                          '<Default ContentType="image/gif" Extension="gif"/>' + \
                          '<Default ContentType="image/jpeg" Extension="jpeg"/>' + \
                          '<Default ContentType="image/png" Extension="png"/>' + \
                          '<Override PartName="/word/document.xml" ' + \
                          'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' + \
                          '<Override PartName="/word/styles.xml" ' + \
                          'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' + \
                          '<Default Extension="rels" ' + \
                          'ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' + \
                          '<Override PartName="/word/webSettings.xml" ' + \
                          'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml"/>' + \
                          '<Override PartName="/word/theme/theme1.xml" ' + \
                          'ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>' + \
                          '<Override PartName="/docProps/core.xml" ' + \
                          'ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' + \
                          '<Override PartName="/word/settings.xml" ' + \
                          'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>' + \
                          '<Override PartName="/docProps/app.xml" ' + \
                          'ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>' + \
                          '</Types>'

    def _get_xml(self):
        return self.xml_string


class SettingsFile(DocxDocument):

    def __init__(self):
        self.dir_name = 'word'
        self.file_name = 'settings.xml'
        self.content = ''
        self.xml_string = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + \
                           '<w:settings xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:sl="http://schemas.openxmlformats.org/schemaLibrary/2006/main">' + \
                               '<w:zoom w:percent="90" />' + \
                               '<w:embedSystemFonts />' + \
                               '<w:proofState w:spelling="clean" w:grammar="clean" />' + \
                               '<w:doNotTrackMoves />' + \
                               '<w:defaultTabStop w:val="720" />' + \
                               '<w:drawingGridHorizontalSpacing w:val="360" />' + \
                               '<w:drawingGridVerticalSpacing w:val="360" />' + \
                               '<w:displayHorizontalDrawingGridEvery w:val="0" />' + \
                               '<w:displayVerticalDrawingGridEvery w:val="0" />' + \
                               '<w:characterSpacingControl w:val="doNotCompress" />' + \
                               '<w:savePreviewPicture />' + \
                               '<w:compat>' + \
                                   '<w:doNotAutofitConstrainedTables />' + \
                                   '<w:doNotVertAlignCellWithSp />' + \
                                   '<w:doNotBreakConstrainedForcedTable />' + \
                                   '<w:useAnsiKerningPairs />' + \
                                   '<w:cachedColBalance />' + \
                                   '<w:splitPgBreakAndParaMark />' + \
                               '</w:compat>' + \
                               '<w:rsids>' + \
                                   '<w:rsidRoot w:val="00854B22" />' + \
                                   '<w:rsid w:val="00854B22" />' + \
                               '</w:rsids>' + \
                               '<m:mathPr>' + \
                                   '<m:mathFont m:val="Arial Black" />' + \
                                   '<m:brkBin m:val="before" />' + \
                                   '<m:brkBinSub m:val="--" />' + \
                                   '<m:smallFrac m:val="off" />' + \
                                   '<m:dispDef m:val="off" />' + \
                                   '<m:lMargin m:val="0" />' + \
                                   '<m:rMargin m:val="0" />' + \
                                   '<m:wrapRight />' + \
                                   '<m:intLim m:val="subSup" />' + \
                                   '<m:naryLim m:val="subSup" />' + \
                               '</m:mathPr>' + \
                               '<w:themeFontLang w:val="en-US" />' + \
                               '<w:clrSchemeMapping w:bg1="light1" w:t1="dark1" w:bg2="light2" w:t2="dark2" w:accent1="accent1" w:accent2="accent2" w:accent3="accent3" w:accent4="accent4" w:accent5="accent5" w:accent6="accent6" w:hyperlink="hyperlink" w:followedHyperlink="followedHyperlink" />' + \
                               '<w:decimalSymbol w:val="." />' + \
                               '<w:listSeparator w:val="," />' + \
                           '</w:settings>'

    def _get_xml(self):
        return self.xml_string.format(content=self.content)

    def _set_content(self, content):
        self.content = content


class FontTableFile(DocxDocument):

    def __init__(self):
        self.dir_name = 'word'
        self.file_name = 'fontTable.xml'
        self.xml_string = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + \
                          '<w:fonts xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' + \
                              '<w:font w:name="Cambria">' + \
                                  '<w:panose1 w:val="02040503050406030204" />' + \
                                  '<w:charset w:val="00" />' + \
                                  '<w:family w:val="auto" />' + \
                                  '<w:pitch w:val="variable" />' + \
                                  '<w:sig w:usb0="00000003" w:usb1="00000000" w:usb2="00000000" w:usb3="00000000" w:csb0="00000001" w:csb1="00000000" />' + \
                              '</w:font>' + \
                              '<w:font w:name="Times New Roman">' + \
                                  '<w:panose1 w:val="02020603050405020304" />' + \
                                  '<w:charset w:val="00" />' + \
                                  '<w:family w:val="auto" />' + \
                                  '<w:pitch w:val="variable" />' + \
                                  '<w:sig w:usb0="00000003" w:usb1="00000000" w:usb2="00000000" w:usb3="00000000" w:csb0="00000001" w:csb1="00000000" />' + \
                              '</w:font>' + \
                              '<w:font w:name="Arial">' + \
                                  '<w:panose1 w:val="020B0604020202020204" />' + \
                                  '<w:charset w:val="00" />' + \
                                  '<w:family w:val="auto" />' + \
                                  '<w:pitch w:val="variable" />' + \
                                  '<w:sig w:usb0="00000003" w:usb1="00000000" w:usb2="00000000" w:usb3="00000000" w:csb0="00000001" w:csb1="00000000" />' + \
                              '</w:font>' + \
                              '<w:font w:name="Calibri">' + \
                                  '<w:panose1 w:val="020F0502020204030204" />' + \
                                  '<w:charset w:val="00" />' + \
                                  '<w:family w:val="auto" />' + \
                                  '<w:pitch w:val="variable" />' + \
                                  '<w:sig w:usb0="00000003" w:usb1="00000000" w:usb2="00000000" w:usb3="00000000" w:csb0="00000001" w:csb1="00000000" />' + \
                              '</w:font>' + \
                          '</w:fonts>'

    def _get_xml(self):
        return self.xml_string


class WebSettingsFile(DocxDocument):

    def __init__(self):
        self.dir_name = 'word'
        self.file_name = 'webSettings.xml'
        self.xml_string = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + \
                          '<w:webSettings ' + \
                          'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' + \
                          'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' + \
                          '<w:allowPNG/>' + \
                          '<w:doNotSaveAsSingleFile/>' + \
                          '</w:webSettings>'

    def _get_xml(self):
        return self.xml_string


class StyleFile(DocxDocument):

    def __init__(self):
        self.dir_name = 'word'
        self.file_name = 'styles.xml'
        self.xml_string = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + \
                            '<w:styles xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' + \
                              '<w:docDefaults>' + \
                                '<w:rPrDefault>' + \
                                  '<w:rPr>' + \
                                    '<w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:cstheme="minorBidi"/>' + \
                                    '<w:sz w:val="24"/>' + \
                                    '<w:szCs w:val="24"/>' + \
                                    '<w:lang w:val="en-US" w:eastAsia="en-US" w:bidi="ar-SA"/>' + \
                                  '</w:rPr>' + \
                                '</w:rPrDefault>' + \
                                '<w:pPrDefault>' + \
                                  '<w:pPr>' + \
                                    '<w:spacing w:after="200"/>' + \
                                  '</w:pPr>' + \
                                '</w:pPrDefault>' + \
                              '</w:docDefaults>' + \
                              '<w:latentStyles w:defLockedState="0" w:defUIPriority="99" w:defSemiHidden="1" w:defUnhideWhenUsed="1" w:defQFormat="0" w:count="276">' + \
                                '<w:lsdException w:name="Normal" w:semiHidden="0" w:uiPriority="0" w:unhideWhenUsed="0" w:qFormat="1"/>' + \
                                '<w:lsdException w:name="heading 1" w:semiHidden="0" w:uiPriority="9" w:unhideWhenUsed="0" w:qFormat="1"/>' + \
                                '<w:lsdException w:name="heading 2" w:uiPriority="9" w:qFormat="1"/>' + \
                                '<w:lsdException w:name="heading 3" w:uiPriority="9" w:qFormat="1"/>' + \
                                '<w:lsdException w:name="heading 4" w:uiPriority="9" w:qFormat="1"/>' + \
                                '<w:lsdException w:name="heading 5" w:uiPriority="9" w:qFormat="1"/>' + \
                                '<w:lsdException w:name="heading 6" w:uiPriority="9" w:qFormat="1"/>' + \
                                '<w:lsdException w:name="heading 7" w:uiPriority="9" w:qFormat="1"/>' + \
                                '<w:lsdException w:name="heading 8" w:uiPriority="9" w:qFormat="1"/>' + \
                                '<w:lsdException w:name="heading 9" w:uiPriority="9" w:qFormat="1"/>' + \
                                '<w:lsdException w:name="toc 1" w:uiPriority="39"/>' + \
                                '<w:lsdException w:name="toc 2" w:uiPriority="39"/>' + \
                                '<w:lsdException w:name="toc 3" w:uiPriority="39"/>' + \
                                '<w:lsdException w:name="toc 4" w:uiPriority="39"/>' + \
                                '<w:lsdException w:name="toc 5" w:uiPriority="39"/>' + \
                                '<w:lsdException w:name="toc 6" w:uiPriority="39"/>' + \
                                '<w:lsdException w:name="toc 7" w:uiPriority="39"/>' + \
                                '<w:lsdException w:name="toc 8" w:uiPriority="39"/>' + \
                                '<w:lsdException w:name="toc 9" w:uiPriority="39"/>' + \
                                '<w:lsdException w:name="caption" w:uiPriority="35" w:qFormat="1"/>' + \
                                '<w:lsdException w:name="Title" w:semiHidden="0" w:uiPriority="10" w:unhideWhenUsed="0" w:qFormat="1"/>' + \
                                '<w:lsdException w:name="Default Paragraph Font" w:uiPriority="1"/>' + \
                                '<w:lsdException w:name="Subtitle" w:semiHidden="0" w:uiPriority="11" w:unhideWhenUsed="0" w:qFormat="1"/>' + \
                                '<w:lsdException w:name="Strong" w:semiHidden="0" w:uiPriority="22" w:unhideWhenUsed="0" w:qFormat="1"/>' + \
                                '<w:lsdException w:name="Emphasis" w:semiHidden="0" w:uiPriority="20" w:unhideWhenUsed="0" w:qFormat="1"/>' + \
                                '<w:lsdException w:name="Table Grid" w:semiHidden="0" w:uiPriority="59" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Placeholder Text" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="No Spacing" w:semiHidden="0" w:uiPriority="1" w:unhideWhenUsed="0" w:qFormat="1"/>' + \
                                '<w:lsdException w:name="Light Shading" w:semiHidden="0" w:uiPriority="60" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Light List" w:semiHidden="0" w:uiPriority="61" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Light Grid" w:semiHidden="0" w:uiPriority="62" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Medium Shading 1" w:semiHidden="0" w:uiPriority="63" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Medium Shading 2" w:semiHidden="0" w:uiPriority="64" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Medium List 1" w:semiHidden="0" w:uiPriority="65" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Medium List 2" w:semiHidden="0" w:uiPriority="66" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Medium Grid 1" w:semiHidden="0" w:uiPriority="67" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Medium Grid 2" w:semiHidden="0" w:uiPriority="68" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Medium Grid 3" w:semiHidden="0" w:uiPriority="69" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Dark List" w:semiHidden="0" w:uiPriority="70" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Colorful Shading" w:semiHidden="0" w:uiPriority="71" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Colorful List" w:semiHidden="0" w:uiPriority="72" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Colorful Grid" w:semiHidden="0" w:uiPriority="73" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Light Shading Accent 1" w:semiHidden="0" w:uiPriority="60" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Light List Accent 1" w:semiHidden="0" w:uiPriority="61" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Light Grid Accent 1" w:semiHidden="0" w:uiPriority="62" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Medium Shading 1 Accent 1" w:semiHidden="0" w:uiPriority="63" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Medium Shading 2 Accent 1" w:semiHidden="0" w:uiPriority="64" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Medium List 1 Accent 1" w:semiHidden="0" w:uiPriority="65" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Revision" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="List Paragraph" w:semiHidden="0" w:uiPriority="34" w:unhideWhenUsed="0" w:qFormat="1"/>' + \
                                '<w:lsdException w:name="Quote" w:semiHidden="0" w:uiPriority="29" w:unhideWhenUsed="0" w:qFormat="1"/>' + \
                                '<w:lsdException w:name="Intense Quote" w:semiHidden="0" w:uiPriority="30" w:unhideWhenUsed="0" w:qFormat="1"/>' + \
                                '<w:lsdException w:name="Medium List 2 Accent 1" w:semiHidden="0" w:uiPriority="66" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Medium Grid 1 Accent 1" w:semiHidden="0" w:uiPriority="67" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Medium Grid 2 Accent 1" w:semiHidden="0" w:uiPriority="68" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Medium Grid 3 Accent 1" w:semiHidden="0" w:uiPriority="69" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Dark List Accent 1" w:semiHidden="0" w:uiPriority="70" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Colorful Shading Accent 1" w:semiHidden="0" w:uiPriority="71" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Colorful List Accent 1" w:semiHidden="0" w:uiPriority="72" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Colorful Grid Accent 1" w:semiHidden="0" w:uiPriority="73" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Light Shading Accent 2" w:semiHidden="0" w:uiPriority="60" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Light List Accent 2" w:semiHidden="0" w:uiPriority="61" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Light Grid Accent 2" w:semiHidden="0" w:uiPriority="62" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Medium Shading 1 Accent 2" w:semiHidden="0" w:uiPriority="63" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Medium Shading 2 Accent 2" w:semiHidden="0" w:uiPriority="64" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Medium List 1 Accent 2" w:semiHidden="0" w:uiPriority="65" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Medium List 2 Accent 2" w:semiHidden="0" w:uiPriority="66" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Medium Grid 1 Accent 2" w:semiHidden="0" w:uiPriority="67" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Medium Grid 2 Accent 2" w:semiHidden="0" w:uiPriority="68" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Medium Grid 3 Accent 2" w:semiHidden="0" w:uiPriority="69" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Dark List Accent 2" w:semiHidden="0" w:uiPriority="70" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Colorful Shading Accent 2" w:semiHidden="0" w:uiPriority="71" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Colorful List Accent 2" w:semiHidden="0" w:uiPriority="72" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Colorful Grid Accent 2" w:semiHidden="0" w:uiPriority="73" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Light Shading Accent 3" w:semiHidden="0" w:uiPriority="60" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Light List Accent 3" w:semiHidden="0" w:uiPriority="61" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Light Grid Accent 3" w:semiHidden="0" w:uiPriority="62" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Medium Shading 1 Accent 3" w:semiHidden="0" w:uiPriority="63" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Medium Shading 2 Accent 3" w:semiHidden="0" w:uiPriority="64" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Medium List 1 Accent 3" w:semiHidden="0" w:uiPriority="65" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Medium List 2 Accent 3" w:semiHidden="0" w:uiPriority="66" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Medium Grid 1 Accent 3" w:semiHidden="0" w:uiPriority="67" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Medium Grid 2 Accent 3" w:semiHidden="0" w:uiPriority="68" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Medium Grid 3 Accent 3" w:semiHidden="0" w:uiPriority="69" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Dark List Accent 3" w:semiHidden="0" w:uiPriority="70" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Colorful Shading Accent 3" w:semiHidden="0" w:uiPriority="71" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Colorful List Accent 3" w:semiHidden="0" w:uiPriority="72" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Colorful Grid Accent 3" w:semiHidden="0" w:uiPriority="73" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Light Shading Accent 4" w:semiHidden="0" w:uiPriority="60" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Light List Accent 4" w:semiHidden="0" w:uiPriority="61" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Light Grid Accent 4" w:semiHidden="0" w:uiPriority="62" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Medium Shading 1 Accent 4" w:semiHidden="0" w:uiPriority="63" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Medium Shading 2 Accent 4" w:semiHidden="0" w:uiPriority="64" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Medium List 1 Accent 4" w:semiHidden="0" w:uiPriority="65" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Medium List 2 Accent 4" w:semiHidden="0" w:uiPriority="66" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Medium Grid 1 Accent 4" w:semiHidden="0" w:uiPriority="67" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Medium Grid 2 Accent 4" w:semiHidden="0" w:uiPriority="68" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Medium Grid 3 Accent 4" w:semiHidden="0" w:uiPriority="69" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Dark List Accent 4" w:semiHidden="0" w:uiPriority="70" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Colorful Shading Accent 4" w:semiHidden="0" w:uiPriority="71" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Colorful List Accent 4" w:semiHidden="0" w:uiPriority="72" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Colorful Grid Accent 4" w:semiHidden="0" w:uiPriority="73" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Light Shading Accent 5" w:semiHidden="0" w:uiPriority="60" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Light List Accent 5" w:semiHidden="0" w:uiPriority="61" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Light Grid Accent 5" w:semiHidden="0" w:uiPriority="62" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Medium Shading 1 Accent 5" w:semiHidden="0" w:uiPriority="63" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Medium Shading 2 Accent 5" w:semiHidden="0" w:uiPriority="64" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Medium List 1 Accent 5" w:semiHidden="0" w:uiPriority="65" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Medium List 2 Accent 5" w:semiHidden="0" w:uiPriority="66" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Medium Grid 1 Accent 5" w:semiHidden="0" w:uiPriority="67" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Medium Grid 2 Accent 5" w:semiHidden="0" w:uiPriority="68" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Medium Grid 3 Accent 5" w:semiHidden="0" w:uiPriority="69" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Dark List Accent 5" w:semiHidden="0" w:uiPriority="70" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Colorful Shading Accent 5" w:semiHidden="0" w:uiPriority="71" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Colorful List Accent 5" w:semiHidden="0" w:uiPriority="72" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Colorful Grid Accent 5" w:semiHidden="0" w:uiPriority="73" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Light Shading Accent 6" w:semiHidden="0" w:uiPriority="60" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Light List Accent 6" w:semiHidden="0" w:uiPriority="61" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Light Grid Accent 6" w:semiHidden="0" w:uiPriority="62" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Medium Shading 1 Accent 6" w:semiHidden="0" w:uiPriority="63" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Medium Shading 2 Accent 6" w:semiHidden="0" w:uiPriority="64" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Medium List 1 Accent 6" w:semiHidden="0" w:uiPriority="65" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Medium List 2 Accent 6" w:semiHidden="0" w:uiPriority="66" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Medium Grid 1 Accent 6" w:semiHidden="0" w:uiPriority="67" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Medium Grid 2 Accent 6" w:semiHidden="0" w:uiPriority="68" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Medium Grid 3 Accent 6" w:semiHidden="0" w:uiPriority="69" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Dark List Accent 6" w:semiHidden="0" w:uiPriority="70" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Colorful Shading Accent 6" w:semiHidden="0" w:uiPriority="71" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Colorful List Accent 6" w:semiHidden="0" w:uiPriority="72" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Colorful Grid Accent 6" w:semiHidden="0" w:uiPriority="73" w:unhideWhenUsed="0"/>' + \
                                '<w:lsdException w:name="Subtle Emphasis" w:semiHidden="0" w:uiPriority="19" w:unhideWhenUsed="0" w:qFormat="1"/>' + \
                                '<w:lsdException w:name="Intense Emphasis" w:semiHidden="0" w:uiPriority="21" w:unhideWhenUsed="0" w:qFormat="1"/>' + \
                                '<w:lsdException w:name="Subtle Reference" w:semiHidden="0" w:uiPriority="31" w:unhideWhenUsed="0" w:qFormat="1"/>' + \
                                '<w:lsdException w:name="Intense Reference" w:semiHidden="0" w:uiPriority="32" w:unhideWhenUsed="0" w:qFormat="1"/>' + \
                                '<w:lsdException w:name="Book Title" w:semiHidden="0" w:uiPriority="33" w:unhideWhenUsed="0" w:qFormat="1"/>' + \
                                '<w:lsdException w:name="Bibliography" w:uiPriority="37"/>' + \
                                '<w:lsdException w:name="TOC Heading" w:uiPriority="39" w:qFormat="1"/>' + \
                              '</w:latentStyles>' + \
                              '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">' + \
                                '<w:name w:val="Normal"/>' + \
                                '<w:qFormat/>' + \
                                '<w:rsid w:val="001A4D62"/>' + \
                              '</w:style>' + \
                              '<w:style w:type="paragraph" w:styleId="Heading1">' + \
                                '<w:name w:val="heading 1"/>' + \
                                '<w:basedOn w:val="Normal"/>' + \
                                '<w:next w:val="Normal"/>' + \
                                '<w:link w:val="Heading1Char"/>' + \
                                '<w:uiPriority w:val="9"/>' + \
                                '<w:qFormat/>' + \
                                '<w:rsid w:val="000222C4"/>' + \
                                '<w:pPr>' + \
                                  '<w:keepNext/>' + \
                                  '<w:keepLines/>' + \
                                  '<w:spacing w:before="480" w:after="0"/>' + \
                                  '<w:outlineLvl w:val="0"/>' + \
                                '</w:pPr>' + \
                                '<w:rPr>' + \
                                  '<w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/>' + \
                                  '<w:b/>' + \
                                  '<w:bCs/>' + \
                                  '<w:color w:val="345A8A" w:themeColor="accent1" w:themeShade="B5"/>' + \
                                  '<w:sz w:val="32"/>' + \
                                  '<w:szCs w:val="32"/>' + \
                                '</w:rPr>' + \
                              '</w:style>' + \
                              '<w:style w:type="paragraph" w:styleId="Heading2">' + \
                                '<w:name w:val="heading 2"/>' + \
                                '<w:basedOn w:val="Normal"/>' + \
                                '<w:next w:val="Normal"/>' + \
                                '<w:link w:val="Heading2Char"/>' + \
                                '<w:uiPriority w:val="9"/>' + \
                                '<w:unhideWhenUsed/>' + \
                                '<w:qFormat/>' + \
                                '<w:rsid w:val="000222C4"/>' + \
                                '<w:pPr>' + \
                                  '<w:keepNext/>' + \
                                  '<w:keepLines/>' + \
                                  '<w:spacing w:before="200" w:after="0"/>' + \
                                  '<w:outlineLvl w:val="1"/>' + \
                                '</w:pPr>' + \
                                '<w:rPr>' + \
                                  '<w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/>' + \
                                  '<w:b/>' + \
                                  '<w:bCs/>' + \
                                  '<w:color w:val="4F81BD" w:themeColor="accent1"/>' + \
                                  '<w:sz w:val="26"/>' + \
                                  '<w:szCs w:val="26"/>' + \
                                '</w:rPr>' + \
                              '</w:style>' + \
                              '<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont">' + \
                                '<w:name w:val="Default Paragraph Font"/>' + \
                                '<w:semiHidden/>' + \
                                '<w:unhideWhenUsed/>' + \
                              '</w:style>' + \
                              '<w:style w:type="table" w:default="1" w:styleId="TableNormal">' + \
                                '<w:name w:val="Normal Table"/>' + \
                                '<w:semiHidden/>' + \
                                '<w:unhideWhenUsed/>' + \
                                '<w:qFormat/>' + \
                                '<w:tblPr>' + \
                                  '<w:tblInd w:w="0" w:type="dxa"/>' + \
                                  '<w:tblCellMar>' + \
                                    '<w:top w:w="0" w:type="dxa"/>' + \
                                    '<w:left w:w="108" w:type="dxa"/>' + \
                                    '<w:bottom w:w="0" w:type="dxa"/>' + \
                                    '<w:right w:w="108" w:type="dxa"/>' + \
                                  '</w:tblCellMar>' + \
                                '</w:tblPr>' + \
                              '</w:style>' + \
                              '<w:style w:type="numbering" w:default="1" w:styleId="NoList">' + \
                                '<w:name w:val="No List"/>' + \
                                '<w:semiHidden/>' + \
                                '<w:unhideWhenUsed/>' + \
                              '</w:style>' + \
                              '<w:style w:type="character" w:customStyle="1" w:styleId="Heading1Char">' + \
                                '<w:name w:val="Heading 1 Char"/>' + \
                                '<w:basedOn w:val="DefaultParagraphFont"/>' + \
                                '<w:link w:val="Heading1"/>' + \
                                '<w:uiPriority w:val="9"/>' + \
                                '<w:rsid w:val="000222C4"/>' + \
                                '<w:rPr>' + \
                                  '<w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/>' + \
                                  '<w:b/>' + \
                                  '<w:bCs/>' + \
                                  '<w:color w:val="345A8A" w:themeColor="accent1" w:themeShade="B5"/>' + \
                                  '<w:sz w:val="32"/>' + \
                                  '<w:szCs w:val="32"/>' + \
                                '</w:rPr>' + \
                              '</w:style>' + \
                              '<w:style w:type="character" w:customStyle="1" w:styleId="Heading2Char">' + \
                                '<w:name w:val="Heading 2 Char"/>' + \
                                '<w:basedOn w:val="DefaultParagraphFont"/>' + \
                                '<w:link w:val="Heading2"/>' + \
                                '<w:uiPriority w:val="9"/>' + \
                                '<w:rsid w:val="000222C4"/>' + \
                                '<w:rPr>' + \
                                  '<w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/>' + \
                                  '<w:b/>' + \
                                  '<w:bCs/>' + \
                                  '<w:color w:val="4F81BD" w:themeColor="accent1"/>' + \
                                  '<w:sz w:val="26"/>' + \
                                  '<w:szCs w:val="26"/>' + \
                                '</w:rPr>' + \
                              '</w:style>' + \
                            '</w:styles>'

    def _get_xml(self):
        return self.xml_string


class ThemeFile(DocxDocument):
    def __init__(self):
        self.dir_name = 'word/theme'
        self.file_name = 'theme1.xml'
        self.xml_string = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + \
                          '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">' + \
                            '<a:themeElements>' + \
                              '<a:clrScheme name="Office">' + \
                                '<a:dk1>' + \
                                  '<a:sysClr val="windowText" lastClr="000000"/>' + \
                                '</a:dk1>' + \
                                '<a:lt1>' + \
                                  '<a:sysClr val="window" lastClr="FFFFFF"/>' + \
                                '</a:lt1>' + \
                                '<a:dk2>' + \
                                  '<a:srgbClr val="1F497D"/>' + \
                                '</a:dk2>' + \
                                '<a:lt2>' + \
                                  '<a:srgbClr val="EEECE1"/>' + \
                                '</a:lt2>' + \
                                '<a:accent1>' + \
                                  '<a:srgbClr val="4F81BD"/>' + \
                                '</a:accent1>' + \
                                '<a:accent2>' + \
                                  '<a:srgbClr val="C0504D"/>' + \
                                '</a:accent2>' + \
                                '<a:accent3>' + \
                                  '<a:srgbClr val="9BBB59"/>' + \
                                '</a:accent3>' + \
                                '<a:accent4>' + \
                                  '<a:srgbClr val="8064A2"/>' + \
                                '</a:accent4>' + \
                                '<a:accent5>' + \
                                  '<a:srgbClr val="4BACC6"/>' + \
                                '</a:accent5>' + \
                                '<a:accent6>' + \
                                  '<a:srgbClr val="F79646"/>' + \
                                '</a:accent6>' + \
                                '<a:hlink>' + \
                                  '<a:srgbClr val="0000FF"/>' + \
                                '</a:hlink>' + \
                                '<a:folHlink>' + \
                                  '<a:srgbClr val="800080"/>' + \
                                '</a:folHlink>' + \
                              '</a:clrScheme>' + \
                              '<a:fontScheme name="Office">' + \
                                '<a:majorFont>' + \
                                  '<a:latin typeface="Calibri"/>' + \
                                  '<a:ea typeface=""/>' + \
                                  '<a:cs typeface=""/>' + \
                                  '<a:font script="Jpan" typeface="ＭＳ ゴシック"/>' + \
                                  '<a:font script="Hang" typeface="맑은 고딕"/>' + \
                                  '<a:font script="Hans" typeface="宋体"/>' + \
                                  '<a:font script="Hant" typeface="新細明體"/>' + \
                                  '<a:font script="Arab" typeface="Times New Roman"/>' + \
                                  '<a:font script="Hebr" typeface="Times New Roman"/>' + \
                                  '<a:font script="Thai" typeface="Angsana New"/>' + \
                                  '<a:font script="Ethi" typeface="Nyala"/>' + \
                                  '<a:font script="Beng" typeface="Vrinda"/>' + \
                                  '<a:font script="Gujr" typeface="Shruti"/>' + \
                                  '<a:font script="Khmr" typeface="MoolBoran"/>' + \
                                  '<a:font script="Knda" typeface="Tunga"/>' + \
                                  '<a:font script="Guru" typeface="Raavi"/>' + \
                                  '<a:font script="Cans" typeface="Euphemia"/>' + \
                                  '<a:font script="Cher" typeface="Plantagenet Cherokee"/>' + \
                                  '<a:font script="Yiii" typeface="Microsoft Yi Baiti"/>' + \
                                  '<a:font script="Tibt" typeface="Microsoft Himalaya"/>' + \
                                  '<a:font script="Thaa" typeface="MV Boli"/>' + \
                                  '<a:font script="Deva" typeface="Mangal"/>' + \
                                  '<a:font script="Telu" typeface="Gautami"/>' + \
                                  '<a:font script="Taml" typeface="Latha"/>' + \
                                  '<a:font script="Syrc" typeface="Estrangelo Edessa"/>' + \
                                  '<a:font script="Orya" typeface="Kalinga"/>' + \
                                  '<a:font script="Mlym" typeface="Kartika"/>' + \
                                  '<a:font script="Laoo" typeface="DokChampa"/>' + \
                                  '<a:font script="Sinh" typeface="Iskoola Pota"/>' + \
                                  '<a:font script="Mong" typeface="Mongolian Baiti"/>' + \
                                  '<a:font script="Viet" typeface="Times New Roman"/>' + \
                                  '<a:font script="Uigh" typeface="Microsoft Uighur"/>' + \
                                '</a:majorFont>' + \
                                '<a:minorFont>' + \
                                  '<a:latin typeface="Cambria"/>' + \
                                  '<a:ea typeface=""/>' + \
                                  '<a:cs typeface=""/>' + \
                                  '<a:font script="Jpan" typeface="ＭＳ 明朝"/>' + \
                                  '<a:font script="Hang" typeface="맑은 고딕"/>' + \
                                  '<a:font script="Hans" typeface="宋体"/>' + \
                                  '<a:font script="Hant" typeface="新細明體"/>' + \
                                  '<a:font script="Arab" typeface="Arial"/>' + \
                                  '<a:font script="Hebr" typeface="Arial"/>' + \
                                  '<a:font script="Thai" typeface="Cordia New"/>' + \
                                  '<a:font script="Ethi" typeface="Nyala"/>' + \
                                  '<a:font script="Beng" typeface="Vrinda"/>' + \
                                  '<a:font script="Gujr" typeface="Shruti"/>' + \
                                  '<a:font script="Khmr" typeface="DaunPenh"/>' + \
                                  '<a:font script="Knda" typeface="Tunga"/>' + \
                                  '<a:font script="Guru" typeface="Raavi"/>' + \
                                  '<a:font script="Cans" typeface="Euphemia"/>' + \
                                  '<a:font script="Cher" typeface="Plantagenet Cherokee"/>' + \
                                  '<a:font script="Yiii" typeface="Microsoft Yi Baiti"/>' + \
                                  '<a:font script="Tibt" typeface="Microsoft Himalaya"/>' + \
                                  '<a:font script="Thaa" typeface="MV Boli"/>' + \
                                  '<a:font script="Deva" typeface="Mangal"/>' + \
                                  '<a:font script="Telu" typeface="Gautami"/>' + \
                                  '<a:font script="Taml" typeface="Latha"/>' + \
                                  '<a:font script="Syrc" typeface="Estrangelo Edessa"/>' + \
                                  '<a:font script="Orya" typeface="Kalinga"/>' + \
                                  '<a:font script="Mlym" typeface="Kartika"/>' + \
                                  '<a:font script="Laoo" typeface="DokChampa"/>' + \
                                  '<a:font script="Sinh" typeface="Iskoola Pota"/>' + \
                                  '<a:font script="Mong" typeface="Mongolian Baiti"/>' + \
                                  '<a:font script="Viet" typeface="Arial"/>' + \
                                  '<a:font script="Uigh" typeface="Microsoft Uighur"/>' + \
                                '</a:minorFont>' + \
                              '</a:fontScheme>' + \
                              '<a:fmtScheme name="Office">' + \
                                '<a:fillStyleLst>' + \
                                  '<a:solidFill>' + \
                                    '<a:schemeClr val="phClr"/>' + \
                                  '</a:solidFill>' + \
                                  '<a:gradFill rotWithShape="1">' + \
                                    '<a:gsLst>' + \
                                      '<a:gs pos="0">' + \
                                        '<a:schemeClr val="phClr">' + \
                                          '<a:tint val="50000"/>' + \
                                          '<a:satMod val="300000"/>' + \
                                        '</a:schemeClr>' + \
                                      '</a:gs>' + \
                                      '<a:gs pos="35000">' + \
                                        '<a:schemeClr val="phClr">' + \
                                          '<a:tint val="37000"/>' + \
                                          '<a:satMod val="300000"/>' + \
                                        '</a:schemeClr>' + \
                                      '</a:gs>' + \
                                      '<a:gs pos="100000">' + \
                                        '<a:schemeClr val="phClr">' + \
                                          '<a:tint val="15000"/>' + \
                                          '<a:satMod val="350000"/>' + \
                                        '</a:schemeClr>' + \
                                      '</a:gs>' + \
                                    '</a:gsLst>' + \
                                    '<a:lin ang="16200000" scaled="1"/>' + \
                                  '</a:gradFill>' + \
                                  '<a:gradFill rotWithShape="1">' + \
                                    '<a:gsLst>' + \
                                      '<a:gs pos="0">' + \
                                        '<a:schemeClr val="phClr">' + \
                                          '<a:tint val="100000"/>' + \
                                          '<a:shade val="100000"/>' + \
                                          '<a:satMod val="130000"/>' + \
                                        '</a:schemeClr>' + \
                                      '</a:gs>' + \
                                      '<a:gs pos="100000">' + \
                                        '<a:schemeClr val="phClr">' + \
                                          '<a:tint val="50000"/>' + \
                                          '<a:shade val="100000"/>' + \
                                          '<a:satMod val="350000"/>' + \
                                        '</a:schemeClr>' + \
                                      '</a:gs>' + \
                                    '</a:gsLst>' + \
                                    '<a:lin ang="16200000" scaled="0"/>' + \
                                  '</a:gradFill>' + \
                                '</a:fillStyleLst>' + \
                                '<a:lnStyleLst>' + \
                                  '<a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">' + \
                                    '<a:solidFill>' + \
                                      '<a:schemeClr val="phClr">' + \
                                        '<a:shade val="95000"/>' + \
                                        '<a:satMod val="105000"/>' + \
                                      '</a:schemeClr>' + \
                                    '</a:solidFill>' + \
                                    '<a:prstDash val="solid"/>' + \
                                  '</a:ln>' + \
                                  '<a:ln w="25400" cap="flat" cmpd="sng" algn="ctr">' + \
                                    '<a:solidFill>' + \
                                      '<a:schemeClr val="phClr"/>' + \
                                    '</a:solidFill>' + \
                                    '<a:prstDash val="solid"/>' + \
                                  '</a:ln>' + \
                                  '<a:ln w="38100" cap="flat" cmpd="sng" algn="ctr">' + \
                                    '<a:solidFill>' + \
                                      '<a:schemeClr val="phClr"/>' + \
                                    '</a:solidFill>' + \
                                    '<a:prstDash val="solid"/>' + \
                                  '</a:ln>' + \
                                '</a:lnStyleLst>' + \
                                '<a:effectStyleLst>' + \
                                  '<a:effectStyle>' + \
                                    '<a:effectLst>' + \
                                      '<a:outerShdw blurRad="40000" dist="20000" dir="5400000" rotWithShape="0">' + \
                                        '<a:srgbClr val="000000">' + \
                                          '<a:alpha val="38000"/>' + \
                                        '</a:srgbClr>' + \
                                      '</a:outerShdw>' + \
                                    '</a:effectLst>' + \
                                  '</a:effectStyle>' + \
                                  '<a:effectStyle>' + \
                                    '<a:effectLst>' + \
                                      '<a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0">' + \
                                        '<a:srgbClr val="000000">' + \
                                          '<a:alpha val="35000"/>' + \
                                        '</a:srgbClr>' + \
                                      '</a:outerShdw>' + \
                                    '</a:effectLst>' + \
                                  '</a:effectStyle>' + \
                                  '<a:effectStyle>' + \
                                    '<a:effectLst>' + \
                                      '<a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0">' + \
                                        '<a:srgbClr val="000000">' + \
                                          '<a:alpha val="35000"/>' + \
                                        '</a:srgbClr>' + \
                                      '</a:outerShdw>' + \
                                    '</a:effectLst>' + \
                                    '<a:scene3d>' + \
                                      '<a:camera prst="orthographicFront">' + \
                                        '<a:rot lat="0" lon="0" rev="0"/>' + \
                                      '</a:camera>' + \
                                      '<a:lightRig rig="threePt" dir="t">' + \
                                        '<a:rot lat="0" lon="0" rev="1200000"/>' + \
                                      '</a:lightRig>' + \
                                    '</a:scene3d>' + \
                                    '<a:sp3d>' + \
                                      '<a:bevelT w="63500" h="25400"/>' + \
                                    '</a:sp3d>' + \
                                  '</a:effectStyle>' + \
                                '</a:effectStyleLst>' + \
                                '<a:bgFillStyleLst>' + \
                                  '<a:solidFill>' + \
                                    '<a:schemeClr val="phClr"/>' + \
                                  '</a:solidFill>' + \
                                  '<a:gradFill rotWithShape="1">' + \
                                    '<a:gsLst>' + \
                                      '<a:gs pos="0">' + \
                                        '<a:schemeClr val="phClr">' + \
                                          '<a:tint val="40000"/>' + \
                                          '<a:satMod val="350000"/>' + \
                                        '</a:schemeClr>' + \
                                      '</a:gs>' + \
                                      '<a:gs pos="40000">' + \
                                        '<a:schemeClr val="phClr">' + \
                                          '<a:tint val="45000"/>' + \
                                          '<a:shade val="99000"/>' + \
                                          '<a:satMod val="350000"/>' + \
                                        '</a:schemeClr>' + \
                                      '</a:gs>' + \
                                      '<a:gs pos="100000">' + \
                                        '<a:schemeClr val="phClr">' + \
                                          '<a:shade val="20000"/>' + \
                                          '<a:satMod val="255000"/>' + \
                                        '</a:schemeClr>' + \
                                      '</a:gs>' + \
                                    '</a:gsLst>' + \
                                    '<a:path path="circle">' + \
                                      '<a:fillToRect l="50000" t="-80000" r="50000" b="180000"/>' + \
                                    '</a:path>' + \
                                  '</a:gradFill>' + \
                                  '<a:gradFill rotWithShape="1">' + \
                                    '<a:gsLst>' + \
                                      '<a:gs pos="0">' + \
                                        '<a:schemeClr val="phClr">' + \
                                          '<a:tint val="80000"/>' + \
                                          '<a:satMod val="300000"/>' + \
                                        '</a:schemeClr>' + \
                                      '</a:gs>' + \
                                      '<a:gs pos="100000">' + \
                                        '<a:schemeClr val="phClr">' + \
                                          '<a:shade val="30000"/>' + \
                                          '<a:satMod val="200000"/>' + \
                                        '</a:schemeClr>' + \
                                      '</a:gs>' + \
                                    '</a:gsLst>' + \
                                    '<a:path path="circle">' + \
                                      '<a:fillToRect l="50000" t="50000" r="50000" b="50000"/>' + \
                                    '</a:path>' + \
                                  '</a:gradFill>' + \
                                '</a:bgFillStyleLst>' + \
                              '</a:fmtScheme>' + \
                            '</a:themeElements>' + \
                            '<a:objectDefaults>' + \
                              '<a:spDef>' + \
                                '<a:spPr/>' + \
                                '<a:bodyPr/>' + \
                                '<a:lstStyle/>' + \
                                '<a:style>' + \
                                  '<a:lnRef idx="1">' + \
                                    '<a:schemeClr val="accent1"/>' + \
                                  '</a:lnRef>' + \
                                  '<a:fillRef idx="3">' + \
                                    '<a:schemeClr val="accent1"/>' + \
                                  '</a:fillRef>' + \
                                  '<a:effectRef idx="2">' + \
                                    '<a:schemeClr val="accent1"/>' + \
                                  '</a:effectRef>' + \
                                  '<a:fontRef idx="minor">' + \
                                    '<a:schemeClr val="lt1"/>' + \
                                  '</a:fontRef>' + \
                                '</a:style>' + \
                              '</a:spDef>' + \
                              '<a:lnDef>' + \
                                '<a:spPr/>' + \
                                '<a:bodyPr/>' + \
                                '<a:lstStyle/>' + \
                                '<a:style>' + \
                                  '<a:lnRef idx="2">' + \
                                    '<a:schemeClr val="accent1"/>' + \
                                  '</a:lnRef>' + \
                                  '<a:fillRef idx="0">' + \
                                    '<a:schemeClr val="accent1"/>' + \
                                  '</a:fillRef>' + \
                                  '<a:effectRef idx="1">' + \
                                    '<a:schemeClr val="accent1"/>' + \
                                  '</a:effectRef>' + \
                                  '<a:fontRef idx="minor">' + \
                                    '<a:schemeClr val="tx1"/>' + \
                                  '</a:fontRef>' + \
                                '</a:style>' + \
                              '</a:lnDef>' + \
                            '</a:objectDefaults>' + \
                            '<a:extraClrSchemeLst/>' + \
                          '</a:theme>'

    def _get_xml(self):
        return self.xml_string
