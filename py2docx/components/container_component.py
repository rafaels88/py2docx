# coding: utf-8
from py2docx.components.component import Component


class ContainerComponent(Component):

    def __init__(self, *args, **kw):
        Component.__init__(self, *args, **kw)

    def append(self, component):
        self.template.append(component.draw())
