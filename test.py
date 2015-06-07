# coding: utf-8

from py2docx.components import Paragraph, BreakLine, Text, Image
from py2docx.docx import Docx

doc = Docx()

br = BreakLine()
image = Image('/Users/rafael/Pictures/Fogao.jpg').align('center').width('20%').height('20%')
txt = Text('Hello World!').bold().underline().italic().font('Arial').block()

doc.append(txt)
doc.append(image)
print doc.save('./refactoring.docx')
