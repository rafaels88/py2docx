# coding: utf-8
from py2docx.structure_files import StructureFile


class DocumentRelationship(StructureFile):

    def __init__(self):
        self.dir_name = 'word/_rels'
        self.file_name = 'document.xml.rels'
        self.id_counter = 0
        self.relations = []

    def draw(self):
        relations = "".join(self.relations)
        xml_string = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + \
                      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + \
                      '{0}' + \
                      '</Relationships>'
        return xml_string.format(relations)

    def add_hyperlink(self, url):
        element = '<Relationship Id="{id}" Type="{type}" Target="{target}" TargetMode="External"/>'
        type = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'
        rel_id = self._generate_new_id()
        relation = element.format(id=rel_id, type=type, target=url)
        self.relations.append(relation)
        return rel_id

    def add_image(self, filepath):
        element = '<Relationship Id="{id}" Type="{type}" Target="{target}"/>'
        type = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
        rel_id = self._generate_new_id()
        relation = element.format(id=rel_id, type=type, target=filepath)
        self.relations.append(relation)
        return rel_id

    def _generate_new_id(self):
        self.id_counter += 1
        return "rId{0}".format(self.id_counter)
