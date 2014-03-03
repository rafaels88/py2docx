# coding : utf-8
from docx import Docx
from elements import Block
from elements.text import InlineText, Break
from elements.image import Image
from elements.table import Table, Cell

doc = Docx()

# bold = Bool, italic = Bool, underline {'style': dotted/dashed/solid/double, 'color': '#333ddd'}
# uppercase = Bool, color = #FF0000
t1 = InlineText("Teste de Text", bold=True, italic=True,
                underline={'style': 'solid', 'color': '#4433ff'},
                color="#FF0000")
t2 = InlineText("Teste de Text 2")
t3 = InlineText("Novo Teste sdf dsfdsfdsfdsfdsfdsfdsf dsfdsf dsf sdf sdfsdf")

#i1 = Image("img.jpg")

bl = Block()
bl.append(t1)
bl.append(Break())
bl.append(t2)

# align: left, right, center, justify
bl2 = Block(align='center')
bl2.append(t3)

c1 = Cell()
c1.append(bl)

c2 = Cell()
c2.append(bl)

# valign: top, center, bottom
# margin: 2cm, 2in, 2pt (W3C CSS Format)
# bgcolor: #333333, 333333
# width: 5cm, 5pt, 5in, 50%
# nowrap: True/False (default) Obs.: Nao funciona com width
c3 = Cell(valign='top', bgcolor='#334455', width='1cm')
c3.append(bl2)

# margin: 2cm, 2in, 2pt (W3C CSS Format)
t = Table(margin='1cm')
t.add_column([c1, c2])
t.add_column([c1, c3])

doc.append(bl)
doc.append(t)
#doc.append(i1)

doc.save("novo")
