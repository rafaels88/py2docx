# coding: utf-8
import os
from os.path import basename
from PIL import Image as PILImage
from py2docx.document import DOCUMENT_PATH
from py2docx.util import Unit


class Image(object):

    def __init__(self, image_path, document, align=None,
                 width='100%', height='100%'):
        self.image = open(image_path, 'rb')
        self.image_name = basename(self.image.name).replace(" ", '-')
        self.document = document
        self.align = align
        self.width = width
        self.height = height
        self.xml = '<w:p>' + \
                   '<w:pPr>' + \
                   '{properties}' + \
                   '</w:pPr>' + \
                   '<w:r>' + \
                     '<w:drawing>' + \
                         '<wp:inline distT="0" distB="0" distL="0" distR="0">' + \
                             '<wp:extent cx="{width}" cy="{height}" />' + \
                             '<wp:effectExtent l="25400" t="0" r="0" b="0" />' + \
                             '<wp:docPr id="1" name="Picture 0" descr="{image_name}" />' + \
                             '<wp:cNvGraphicFramePr>' + \
                                 '<a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1" />' + \
                             '</wp:cNvGraphicFramePr>' + \
                             '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' + \
                                 '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' + \
                                     '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">' + \
                                         '<pic:nvPicPr>' + \
                                             '<pic:cNvPr id="0" name="{image_name}" />' + \
                                             '<pic:cNvPicPr />' + \
                                         '</pic:nvPicPr>' + \
                                         '<pic:blipFill>' + \
                                             '<a:blip r:embed="{rel_id}" />' + \
                                             '<a:stretch>' + \
                                                 '<a:fillRect />' + \
                                             '</a:stretch>' + \
                                         '</pic:blipFill>' + \
                                         '<pic:spPr>' + \
                                             '<a:xfrm>' + \
                                                 '<a:off x="0" y="0" />' + \
                                                 '<a:ext cx="{width}" cy="{height}" />' + \
                                             '</a:xfrm>' + \
                                             '<a:prstGeom prst="rect">' + \
                                                 '<a:avLst />' + \
                                             '</a:prstGeom>' + \
                                         '</pic:spPr>' + \
                                     '</pic:pic>' + \
                                 '</a:graphicData>' + \
                             '</a:graphic>' + \
                         '</wp:inline>' + \
                     '</w:drawing>' + \
                   '</w:r>' + \
                   '</w:p>'

        self.xml_props = []
        self._upload_image()
        self._set_properties()

    def _get_image(self):
        return self.image

    def _set_relashionship(self, rel_id):
        self.xml = self.xml.format(rel_id=rel_id)

    def _upload_image(self):
        dir_media = "{0}/word/media".format(DOCUMENT_PATH)
        if not os.path.exists(dir_media):
            os.makedirs(dir_media)
        img_uploaded = open("{0}/{1}".format(dir_media, self.image_name), 'wb')
        img_uploaded.write(self.image.read())
        img_uploaded.close()

    def _set_properties(self):
        self._set_align()
        self.xml = self.xml.replace('{properties}',
                                    ''.join(self.xml_props))
        image_pil = PILImage.open(self.image.name)
        width = Unit.pixel_to_emu(image_pil.size[0])
        height = Unit.pixel_to_emu(image_pil.size[1])

        width_percentage_num = float(self.width[:-1])
        height_percentage_num = float(self.height[:-1])

        width = (width_percentage_num / 100) * width
        height = (height_percentage_num / 100) * height

        self.xml = self.xml.replace("{width}", str(int(width))) \
                           .replace("{height}", str(int(height))) \
                           .replace("{image_name}", self.image_name)

    def _set_align(self):
        if self.align and \
           self.align in ['left',
                          'right',
                          'center',
                          'justify']:

            xml = '<w:jc w:val="{align}"/>'
            self.xml_props.append(xml.format(align=self.align))

    def _get_xml(self):
        rel_id = self.document \
                     .document_relationship_file \
                     ._add_image(self.image_name)
        self._set_relashionship(rel_id)
        return self.xml
