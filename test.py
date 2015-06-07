# coding: utf-8

from py2docx.components import LineBreak, Text, Image, PageBreak
from py2docx.docx import Docx

doc = Docx()

br = LineBreak()
pb = PageBreak()
image = Image('/Users/rafael/Pictures/Fogao.jpg').align('center').width('20%').height('20%')
txt = Text('Hello World!').bold().underline().italic().font('Arial').block()

doc.append(txt)
doc.append(br)
doc.append(txt)
doc.append(pb)
doc.append(image)
print doc.save('./refactoring.docx')
