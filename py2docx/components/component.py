# coding: utf-8
from py2docx.templates.component_template import ComponentTemplate


class Component(object):
    def __init__(self, *args, **kw):
        self.template = ComponentTemplate()

    def draw(self):
        return self.template.render()
