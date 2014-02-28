# coding: utf-8
import os
from document import DOCUMENT_PATH


class Image(object):

    def __init__(self, image_path):
        self.image = open(image_path, 'r')
        self.xml = """
            <ns0:p>
              <ns0:r>
                <ns0:drawing>
                  <ns0:inline xmlns:ns0="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" distT="0" distR="0" distL="0" distB="0">
                    <ns0:extent cy="927100" cx="2895600"/>
                    <ns0:effectExtent r="0" b="0" l="25400" t="0"/>
                    <ns0:docPr id="2" descr="" name="Picture 1"/>
                    <ns0:cNvGraphicFramePr>
                      <ns0:graphicFrameLocks xmlns:ns0="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/>
                    </ns0:cNvGraphicFramePr>
                    <ns0:graphic xmlns:ns0="http://schemas.openxmlformats.org/drawingml/2006/main">
                      <ns0:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                        <ns0:pic xmlns:ns0="http://schemas.openxmlformats.org/drawingml/2006/picture">
                          <ns0:nvPicPr>
                            <ns0:cNvPr id="0" descr="" name="Picture 1"/>
                            <ns0:cNvPicPr>
                              <ns0:picLocks xmlns:ns0="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeArrowheads="1" noChangeAspect="1"/>
                            </ns0:cNvPicPr>
                          </ns0:nvPicPr>
                          <ns0:blipFill>
                            <ns0:blip xmlns:ns0="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:ns1="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ns1:embed="{rel_id}"/>
                            <ns0:srcRect xmlns:ns0="http://schemas.openxmlformats.org/drawingml/2006/main"/>
                            <ns0:stretch xmlns:ns0="http://schemas.openxmlformats.org/drawingml/2006/main">
                              <ns0:fillRect/>
                            </ns0:stretch>
                          </ns0:blipFill>
                          <ns0:spPr bwMode="auto">
                            <ns0:xfrm xmlns:ns0="http://schemas.openxmlformats.org/drawingml/2006/main" flipV="false" rot="0" flipH="false">
                              <ns0:off y="0" x="0"/>
                              <ns0:ext cy="927100" cx="2895600"/>
                            </ns0:xfrm>
                            <ns0:prstGeom xmlns:ns0="http://schemas.openxmlformats.org/drawingml/2006/main" prst="rect">
                              <ns0:avLst/>
                            </ns0:prstGeom>
                          </ns0:spPr>
                        </ns0:pic>
                      </ns0:graphicData>
                    </ns0:graphic>
                  </ns0:inline>
                </ns0:drawing>
              </ns0:r>
            </ns0:p>
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
