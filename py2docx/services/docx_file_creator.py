# coding: utf-8
import zipfile
import os
import shutil
from py2docx.config import DOCUMENT_PATH
from py2docx.services.structure_file_creator import StructureFileCreator
from py2docx.structure_files.app import App
from py2docx.structure_files.content_type import ContentType
from py2docx.structure_files.core import Core
from py2docx.structure_files.document import Document
from py2docx.structure_files.document_relationship import DocumentRelationship
from py2docx.structure_files.font_table import FontTable
from py2docx.structure_files.relationship import Relationship
from py2docx.structure_files.settings import Settings
from py2docx.structure_files.style import Style
from py2docx.structure_files.theme import Theme
from py2docx.structure_files.web_settings import WebSettings


class DocxFileCreator(object):

    def __init__(self, document):
        self.document = document
        self.file_creator = StructureFileCreator()

    def create(self, path):
        self._create_structure_files()
        self._create_docx_file(path)
        self._clean_temp_files()
        self._create_initial_temp_structure()

    def _clean_temp_files(self):
        if os.path.exists(DOCUMENT_PATH):
            shutil.rmtree(DOCUMENT_PATH)

    def _create_initial_temp_structure(self):
        os.makedirs(DOCUMENT_PATH)
        obj_file = open("{0}/__init__.py".format(DOCUMENT_PATH), 'w')
        obj_file.close()

    def _create_docx_file(self, path):
        zip_name = zipfile.ZipFile("{0}".format(path), 'w')

        for dirpath, dirs, files in os.walk("{0}".format(DOCUMENT_PATH)):
            for f in files:
                if f != '__init__.py':
                    file_name = os.path.join(dirpath, f)
                    file_name_zip = file_name.replace('{0}'
                                                      .format(DOCUMENT_PATH),
                                                      '')
                    zip_name.write(file_name, file_name_zip)

    def _create_structure_files(self):
        structure_files = [
            self._build_relationship(),
            self._build_document_relationship(),
            self._build_app(),
            self._build_core(),
            self._build_content_type(),
            self._build_settings(),
            self._build_fonttable(),
            self._build_websettings(),
            self._build_theme(),
            self._build_style(),
            self._build_document()
        ]
        for obj in structure_files:
            self.file_creator.create(obj)

    def _build_relationship(self):
        return Relationship()

    def _build_document_relationship(self):
        self._document_relationship = DocumentRelationship()
        return self._document_relationship

    def _build_app(self):
        return App()

    def _build_core(self):
        return Core()

    def _build_document(self):
        structure_file = Document()
        images = self.document.retrieve_images()
        for image in images:
            image_path = image.image_path()
            relation_id = self._document_relationship.add_image(image_path)
            image.set_relationship_id(relation_id)
        structure_file.set_document_component(self.document)
        return structure_file

    def _build_content_type(self):
        return ContentType()

    def _build_settings(self):
        return Settings()

    def _build_fonttable(self):
        return FontTable()

    def _build_websettings(self):
        return WebSettings()

    def _build_theme(self):
        return Theme()

    def _build_style(self):
        return Style()
