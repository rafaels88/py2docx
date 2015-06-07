# coding: utf-8
import os
from py2docx.config import DOCUMENT_PATH


class StructureFileCreator(object):

    def create(self, structure_file):
        dir_name = structure_file.dir_name
        file_name = structure_file.file_name
        content = structure_file.draw()

        dir_path = self._create_dir_if_unexistent(dir_name)
        obj_file = open("{0}{1}".format(dir_path, file_name), 'w')
        obj_file.write(content)
        obj_file.close()

    def _create_dir_if_unexistent(self, dir_name):
        dir_path = "{0}/{1}/".format(DOCUMENT_PATH,
                                     dir_name.rstrip('/'))
        if not os.path.exists(dir_path):
            os.makedirs(dir_path)
        return dir_path
