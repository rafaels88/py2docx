# coding: utf-8
import os
from nose.tools import assert_in
from lettuce import step, world, before, after
from py2docx.docx import Docx


@before.all
def setup_document_config():
    world.filename = 'test.docx'
    world.xml_paragraph = '<w:p>' + \
                          '<w:pPr>' + \
                          '{properties}' + \
                          '</w:pPr>' + \
                          '{content}' + \
                          '</w:p>'


@after.all
def delete_document_file(total):
    try:
        os.remove(world.filename)
    except OSError:
        pass


@step
def have_a_docx_document_instance(step):
    world.document = Docx()


@step
def get_a_new_paragraph_instance(step):
    world.paragraph = world.document.new_paragraph()


@step
def append_the_new_paragraph_instance(step):
    world.document.append(world.paragraph)


@step
def a_blank_paragraph_is_added(step):
    world.document.save(world.filename)
    docx_file = open(world.filename, 'r')
    content = docx_file.read()
    expected = world.xml_paragraph.format(properties='',
                                          content='')
    assert_in(expected, content), \
        "Got {0}".format(content)
