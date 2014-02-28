# coding : utf-8
from docx import Docx
from elements.text import Paragraph, InlineText, Break
from elements.image import Image
from elements.table import Table, Cell

doc = Docx()
t1 = InlineText("Teste de Text")
t2 = InlineText("Teste de Text 2")

#i1 = Image("img.jpg")

p1 = Paragraph()
p1.append(t1)
p1.append(Break())
p1.append(t2)


c1 = Cell(bgcolor='#333333', margin='20px')
c1.append(p1)

c2 = Cell()
c2.append(p1)

t = Table()
t.add_column([c1, c2])
t.add_column([c1])

doc.append(p1)
doc.append(t)
#doc.append(i1)

doc.save("novo")
