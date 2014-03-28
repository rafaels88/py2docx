Py2Docx - 0.1.1v (Release Date: March 28, 2014)
=======

Py2Docx is a python module to write .docx documents (>= Microsoft Word 2007).

# Last Modifications
- Image element needs a new argument (document) on its constructor
- Fix on Image insertion on Cell element

# Instalation
```
$ pip install py2docx
```


# Example

```python
# coding: utf-8
from py2docx.docx import Docx
from py2docx.elements import Block
from py2docx.elements.text import InlineText, Break, BlockText
from py2docx.elements.image import Image
from py2docx.elements.table import Table, Cell

doc = Docx()
t1 = InlineText("An Inline Text", bold=True, italic=True,
                underline={'style': 'solid', 'color': '#4433ff'},
                color="#FF0000")

bl = Block()
bl.append(t1)

bl2 = Block(align='center')
bl2.append(t1)

t = Table(width="100%", padding='5pt', border={'left': {'style': 'dashed'}})
c1 = Cell()
c1.append(bl)
c2 = Cell([bl])
c3 = Cell(bl2, valign='top', border={'left': {'size': '1pt', 'color': '#33ddff'}})

t.add_row([c1, c2])
t.add_row([c1, c3])

bl3 = Block([InlineText("Inline Text "),
             InlineText("Bold Here", bold=True)])

i = Image("Pictures/into_the_wild.jpg", document=doc, align='center')

doc.append(bl)
doc.append(t)
doc.append(Block(Break()))
doc.append(i)
doc.append(BlockText("This is a Block Text"))
doc.append(Block(InlineText("This is a Block Text")))
doc.append(bl3)
doc.append(BlockText(u"Arial Test", font='Arial', size=20))

doc.save("./py2docx.docx")
```


# API

## Creating a Document

### Docx
```python
from py2docx.docx import Docx
```

Create a document

###### Methods:
##### append(elem)

Parameter | Description
--------- | -----------
elem      | An Image, Block or Table to put in the document.

```python
doc = Docx()
bl = Block()
doc.append(bl)
```

##### save(path)

Parameter | Description
--------- | -----------
path      | The path that the document are going to be saved.

```python
doc = Docx()
doc.save("./example.docx")
```


## Elements

### Block
```python
from py2docx.docx.elements import Block
```

Block elements is almost the same like "\<div\>" HTML element.

###### Methods:
##### __init__(initial=None, align=None)

Parameter | Description
--------- | -----------
initial | One element or a list of elements to put inside the block.
align | Horizontal align. Options are: 'left', 'right', 'center' or 'justify'.

```python
text = InlineText("Hello World!")
block = Block(text, align='center')
```

##### append(elem)

```python
text = InlineText("Hello World!")
block = Block()
block.append(text)
```

Parameter  | Description
---------- | -----------
elem | Any element to put inside the block.



### Image
```python
from py2docx.elements.image import Image
```

###### Methods:
##### __init__(path, document, align=None)

The accepted types are: png, jpg, gif, jpeg.

Parameters | Description
---------- | -----------
path       | A string with the image's path.
document   | The instance of the document (Docx())
align      | Horizontal Align. Values should be: 'left', 'center' or 'right'

```python
doc = Docx()
Image("/Pictures/image.png", document=doc, align='right')
```



### Table
```python
from py2docx.elements.table import Table
```

###### Methods:
##### __init__(padding=None, width=None, border=None)

Parameters | Description
---------- | -----------
padding    | Padding for all cells. Should be in one of these units: cm (centimeters), in (inches) or pt (points). The numbers should be in the W3C CSS Format.
width      | Width of the table. Should be in one of these units: % (percentage), cm (centimeters), in (inches) or pt (points).
border     | A dict with the specifications. Should be in this format: {'[SIDE]': {'color': '#[HEX]', 'size': '[INT]pt', style: '[dotted,dashed,solid,double]'}. The maximum size of the border is 12pt, minimum is 0.5pt.

```python
Table(width='100%', padding='2cm', border={'left': {'color': '#FF0000', 'size': '2pt', style: 'dotted',
                                           'bottom': {'color': '#FF0000', 'size': '2pt', style: 'dashed',
                                           'top': {'color': '#FFFFFF', 'size': '3pt', style: 'solid',
                                           'right': {'color': '#000000', 'size': '3pt', style: 'double'})
```

##### add_row(cells)

Parameters | Description
---------- | -----------
cells      | A list of Cells to put on a row.

```python
t = Table(width='100%')
t.add_row([Cell(), Cell()])
```



### Cell
```python
from py2docx.elements.table import Cell
```

###### Methods:
##### __init__(initial=None, bgcolor=None, padding=None, width=None, valign=None, nowrap=None, border=None, colspan=1)


Parameters | Description
---------- | -----------
initial    | One element or a list of elements to put inside the cell.
bgcolor    | Background color of the cell, in hexadecimal '#00FF66'.
padding    | Padding for cell. Should be in one of these units: cm (centimeters), in (inches) or pt (points). The numbers should be in the W3C CSS Format.
width      | Width of the cell. Should be in one of these units: % (percentage), cm (centimeters), in (inches) or pt (points).
valign     | Vertical Align. Options are: 'top', 'center', 'bottom'.
nowrap     | True or False. It does not work with width.
border     | A dict with the specifications. Should be in this format: {'[SIDE]': {'color': '#[HEX]', 'size': '[INT]pt', style: '[dotted,dashed,solid,double]'}. The maximum size of the border is 12pt, minimum is 0.5pt.
colspan    | An int with the number of cells.

```python
Cell([Image("path/filename.ext"), BlockText("Hello World!")], bgcolor='#3377FF', padding='5cm 10cm',
     width='5cm', valign='center', border={'bottom': {'color': '#FF0000', 'size': '2pt', style: 'dashed'}, colspan=2)
```


##### append(elem)

```python
text = InlineText("Hello World!")
c = Cell()
c.append(text)
```

Parameter  | Description
---------- | -----------
elem       | A Cell, Block, BlockText or Image, to put inside the cell.



### InlineText

```python
from py2docx.elements.text import InlineText
```

You should put this in a Block. This is like a "\<span\>" HTML element.

###### Methods:
##### __init__(text, bold=None, italic=None, underline=None, uppercase=None, color=None, font=None, size=None)


Parameters | Description
---------- | -----------
text       | A string with words.
bold       | True or False.
italic     | True or False.
underline  | A dict with properties: {'style': '[dotted/dashed/solid/double]', 'color': '#[HEX]'}.
uppercase  | True or False.
color      | Hexadecimal color.
font       | Should be 'Cambria', 'Times New Roman', 'Arial' or 'Calibri'.
size       | An INT of font's size in point.

```python
InlineText("Hello World!", bold=True, italic=True,
                underline={'style': 'solid', 'color': '#4433ff'},
                uppercase=True, color="#FF0000", font="Times New Roman",
                size=14)
```



### Break
```python
from py2docx.elements.text import Break
```

The same as "\<br\ \>" HTML element.

```python
hello = InlineText("Hello")
world = InlineText("World", bold=True)

Block([hello, Break(), world])
```



### BlockText
```python
from py2docx.elements.text import BlockText
```

The same thing as:

```python
Block(InlineText("Hello World"))
```

###### Methods:
##### __init__(text, bold=None, italic=None, underline=None, uppercase=None, color=None, font=None, size=None)


Parameters | Description
---------- | -----------
text       | A string with words.
bold       | True or False.
italic     | True or False.
underline  | A dict with properties: {'style': '[dotted/dashed/solid/double]', 'color': '#[HEX]'}.
uppercase  | True or False.
color      | Hexadecimal color.
font       | Should be 'Cambria', 'Times New Roman', 'Arial' or 'Calibri'.
size       | An INT of font's size in point.

```python
BlockText("Hello World!", bold=True, italic=True,
          underline={'style': 'solid', 'color': '#4433ff'},
          uppercase=True, color="#FF0000", font="Times New Roman",
          size=14)
```



# For Devs

This module needs help. Now, I am writing the unit tests. 
I have made py2docx because there is no good module for build a docx with Python. 
I really needed to be fast. If you want to join me on this project, I will be very very glad.

Download the project, install the requirements and give some ideas to improve this a lot =)

```
make requirements
make unit
```
