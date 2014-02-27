# coding: utf-8


class Image(object):

    def __init__(self, image_path):
        self.image = open(image_path, 'r')
        self.xml = """
            <w:drawing>
                <wp:inline distT="0" distB="0" distL="0" distR="0">
                    <wp:extent cx="{width}" cy="{height}"/>
                    <wp:effectExtent l="19050" t="0" r="0" b="0"/>
                    <wp:docPr id="1" name="{image_name}" descr="{image_name}"/>
                    <wp:cNvGraphicFramePr>
                        <a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/>
                    </wp:cNvGraphicFramePr>
                    <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                            <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
                                <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
                                    <pic:nvPicPr>
                                        <pic:cNvPr id="0" name="{image_name}"/>
                                        <pic:cNvPicPr/>
                                    </pic:nvPicPr>
                                    <pic:blipFill>
                                        <a:blip r:embed="{rel_id}" cstate="print"/>
                                        <a:stretch>
                                            <a:fillRect/>
                                        </a:stretch/>
                                    </pic:blipFill>
                                    <pic:spPr>
                                        <a:xfrm>
                                            <a:off x="0" y="0"/>
                                            <a:ext cx="{width}" cy="{height}"/>
                                        </a:xfrm>
                                        <a:prstGeom rst="rect>
                                            <a:avLst/>
                                        </a:prstGeom>
                                    </pic:spPr>
                                </pic:pic>
                            </pic:pic>
                        </a:graphicData>
                    </a:graphic>
                </wp:inline>
            </w:drawing>
        """

    def _get_xml(self):
        pass
