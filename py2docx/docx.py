# coding: utf-8
import os
import zipfile
from elements.image import Image
from document import RelationshipFile, AppFile, CoreFile, \
    DocumentFile, ContentTypeFile, TMP_PATH, DOCUMENT_PATH, \
    DocumentRelationshipFile, DocxDocument


class Docx(object):

    def __init__(self):
        DocxDocument._clean_all()
        self.content = ''
        self.relationship_file = RelationshipFile()
        self.document_relationship_file = DocumentRelationshipFile()
        self.app_file = AppFile()
        self.core_file = CoreFile()
        self.document_file = DocumentFile()
        self.content_type_file = ContentTypeFile()

    def _create_structure(self):
        self.relationship_file._create_file()
        self.document_relationship_file._create_file()
        self.app_file._create_file()
        self.core_file._create_file()
        self.document_file._create_file()
        self.content_type_file._create_file()

    def _create_document(self, path):
        zip_name = zipfile.ZipFile("{0}".format(path), 'w')

        for dirpath, dirs, files in os.walk("{0}".format(DOCUMENT_PATH)):
            for f in files:
                file_name = os.path.join(dirpath, f)
                file_name_zip = file_name.replace('{0}'
                                                  .format(DOCUMENT_PATH), '')
                zip_name.write(file_name, file_name_zip)

    def save(self, path):
        self.document_file._set_content(self.content)
        self._create_structure()
        self._create_document(path)
        DocxDocument._clean_all()

    def append(self, elem):
        if isinstance(elem, Image):
            rel_id = self._add_image_relationship(elem)
            elem._set_relashionship(rel_id)

        xml_elem = elem._get_xml()
        self.content += xml_elem

    def _add_image_relationship(self, image):
        img_obj = image._get_image()
        rel_id = self.document_relationship_file._add_image(img_obj.name)
        return rel_id
