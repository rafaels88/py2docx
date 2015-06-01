# coding: utf-8
import os
from lettuce import step, world, before, after
from py2docx.docx import Docx


@before.all
def setup_document_config():
    world.filename = 'test.docx'


@after.all
def delete_document_file(total):
    try:
        os.remove(world.filename)
    except OSError:
        pass


@step
def have_a_blank_document(step):
    world.document = Docx()


@step
def save_the_blank_document(step):
    world.document.save(world.filename)


@step
def docx_file_is_created(step):
    docx_file = open(world.filename, 'r')
    content = docx_file.read()
    assert content != '', \
        "Got an empty content"
