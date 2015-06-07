# coding: utf-8
from py2docx.structure_files import StructureFile


class ContentType(StructureFile):

    def __init__(self):
        self.dir_name = ''
        self.file_name = '[Content_Types].xml'

    def draw(self):
        xml_string = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + \
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
        return xml_string
