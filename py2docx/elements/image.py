# coding: utf-8
import os
from document import DOCUMENT_PATH


class Image(object):

    def __init__(self, image_path):
        self.image = open(image_path, 'r')
        self.xml = """
            <w:p>
              <w:r>
                <w:drawing>
                    <wp:inline distT="0" distB="0" distL="0" distR="0">
                        <wp:extent cx="5080000" cy="3810000" />
                        <wp:effectExtent l="25400" t="0" r="0" b="0" />
                        <wp:docPr id="1" name="Picture 0" descr="img.jpg" />
                        <wp:cNvGraphicFramePr>
                            <a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1" />
                        </wp:cNvGraphicFramePr>
                        <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                            <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                                <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
                                    <pic:nvPicPr>
                                        <pic:cNvPr id="0" name="img.jpg" />
                                        <pic:cNvPicPr />
                                    </pic:nvPicPr>
                                    <pic:blipFill>
                                        <a:blip r:embed="{rel_id}" />
                                        <a:stretch>
                                            <a:fillRect />
                                        </a:stretch>
                                    </pic:blipFill>
                                    <pic:spPr>
                                        <a:xfrm>
                                            <a:off x="0" y="0" />
                                            <a:ext cx="5080000" cy="3810000" />
                                        </a:xfrm>
                                        <a:prstGeom prst="rect">
                                            <a:avLst />
                                        </a:prstGeom>
                                    </pic:spPr>
                                </pic:pic>
                            </a:graphicData>
                        </a:graphic>
                    </wp:inline>
                </w:drawing>
              </w:r>
            </w:p>
        """
        self._set_properties()
        self._upload_image()

    def _get_image(self):
        return self.image

    def _set_relashionship(self, rel_id):
        self.xml = self.xml.format(rel_id=rel_id)

    def _upload_image(self):
        dir_media = "{0}/word/media".format(DOCUMENT_PATH)
        if not os.path.exists(dir_media):
            os.makedirs(dir_media)
        file_name = self.image.name
        img_uploaded = open("{0}/{1}".format(dir_media, file_name), 'w')
        img_uploaded.write(self.image.read())
        img_uploaded.close()

    def _set_properties(self):
        name = self.image.name
        width = '100000'
        height = '1000000'
        self.xml = self.xml.replace("{width}", width) \
                           .replace("{height}", height) \
                           .replace("{image_name}", name)

    def _get_xml(self):
        return self.xml
