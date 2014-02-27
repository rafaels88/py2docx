# coding: utf-8


class Paragraph(object):

    def __init__(self, text):
        self.text = text
        self.xml = """
            <ns0:p>
              <ns0:r>
              <ns0:t>{text}</ns0:t>
              </ns0:r>
            </ns0:p>
        """

    def _get_xml(self):
        return self.xml.replace("{text}", self.text)
