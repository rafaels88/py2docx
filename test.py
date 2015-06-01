# coding: utf-8

from py2docx.components import Paragraph, BreakLine, Text
from py2docx.docx import Docx

doc = Docx()

par = Paragraph()
br = BreakLine()
txt = Text('Hello World!').bold().underline().italic().font('Arial')

doc.append(par).append(br).append(txt)
print doc.save()
