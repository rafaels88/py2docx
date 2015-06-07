# coding: utf-8
from py2docx.templates.component_template import ComponentTemplate


class ImageTemplate(ComponentTemplate):

    def __init__(self, image_name, *args, **kw):
        super(ImageTemplate, self).__init__(*args, **kw)
        self.image_name = image_name
        self.properties_list = []
        self.align_property = ''
        self.width_property = ''
        self.height_property = ''
        self.relationship_id = ''

    def begin(self):
        return '<w:p>'

    def properties(self):
        self._build_properties_list()
        xml = '<w:pPr>'
        xml += ''.join(self.properties_list)
        xml += '</w:pPr>'
        return xml

    def contents(self):
        result = '' + \
            '<w:r>' + \
              '<w:drawing>' + \
                '<wp:inline distT="0" distB="0" distL="0" distR="0">' + \
                      '<wp:extent cx="{width}" cy="{height}" />' + \
                      '<wp:effectExtent l="25400" t="0" r="0" b="0" />' + \
                      '<wp:docPr id="1" name="{image_name}" descr="{image_name}" />' + \
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
            '</w:r>'
        result = self._put_values_on_content(result)
        return result

    def end(self):
        return '</w:p>'

    def align(self, value):
        self.align_property = '<w:jc w:val="{0}"/>'.format(value)

    def width(self, value):
        self.width_property = str(value)

    def height(self, value):
        self.height_property = str(value)

    def set_relationship_id(self, id):
        self.relationship_id = str(id)

    def _build_properties_list(self):
        self.properties_list.append(self.align_property)

    def _put_values_on_content(self, content):
        return content.replace('{width}', self.width_property) \
                      .replace('{height}', self.height_property) \
                      .replace('{image_name}', self.image_name) \
                      .replace('{rel_id}', self.relationship_id)
