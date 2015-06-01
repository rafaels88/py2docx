# coding: utf-8


class ComponentTemplate(object):
    def __init__(self):
        self.content_list = []

    def begin(self):
        pass

    def properties(self):
        pass

    def contents(self):
        return ''.join(self.content_list)

    def end(self):
        pass

    def append(self, component):
        self.content_list.append(component)

    def render(self):
        result = self.begin()
        result += self.properties()
        result += self.contents()
        result += self.end()
        return result
