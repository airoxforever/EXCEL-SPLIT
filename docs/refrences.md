indexmodules |next |previous |openpyxl 3.1.3 documentation » openpyxl package » openpyxl.cell package » openpyxl.cell.rich_text module
openpyxl.cell.rich_text module
RichText definition

class openpyxl.cell.rich_text.CellRichText(*args)[source]
Bases: list

Represents a rich text string.

Initialize with a list made of pure strings or TextBlock elements Can index object to access or modify individual rich text elements it also supports the + and += operators between rich text strings There are no user methods for this class

operations which modify the string will generally call an optimization pass afterwards, that merges text blocks with identical formats, consecutive pure text strings, and remove empty strings and empty text blocks

append(arg)[source]
Append object to the end of the list.

as_list()[source]
Returns a list of the strings contained. The main reason for this is to make editing easier.

extend(arg)[source]
Extend list by appending elements from the iterable.

classmethod from_tree(node)[source]
to_tree()[source]
Return the full XML representation

class openpyxl.cell.rich_text.TextBlock(font, text)[source]
Bases: Strict

Represents text string in a specific format

This class is used as part of constructing a rich text strings.

font
Values must be of type <class ‘openpyxl.cell.text.InlineFont’>

text
Values must be of type <class ‘str’>

to_tree()[source]
Logo

Table of Contents
openpyxl.cell.rich_text module
CellRichText
CellRichText.append()
CellRichText.as_list()
CellRichText.extend()
CellRichText.from_tree()
CellRichText.to_tree()
TextBlock
TextBlock.font
TextBlock.text
TextBlock.to_tree()
Previous topic
openpyxl.cell.read_only module

Next topic
openpyxl.cell.text module

---
indexmodules |next |previous |openpyxl 3.1.3 documentation » Working with Rich Text
Working with Rich Text
Introduction
Normally styles apply to everything in an individual cell. However, rich text allows formatting of parts of the text in a string. This section covers adding rich-text formatting to worksheet cells. Rich-text formatting in existing workbooks has to be enabled when loading them with the rich_text=True parameter.

Rich Text objects can contain a mix of unformatted text and TextBlock objects that contains an InlineFont style and a the text which is to be formatted like this. The result is a CellRichText object.

from openpyxl.cell.text import InlineFont
from openpyxl.cell.rich_text import TextBlock, CellRichText
rich_string1 = CellRichText(
   'This is a test ',
   TextBlock(InlineFont(b=True), 'xxx'),
  'yyy'
)
InlineFont objects are virtually identical to the Font objects, but use a different attribute name, rFont, for the name of the font. Unfortunately, this is required by OOXML and cannot be avoided.

inline_font = InlineFont(rFont='Calibri', # Font name
                         sz=22,           # in 1/144 in. (1/2 point) units, must be integer
                         charset=None,    # character set (0 to 255), less required with UTF-8
                         family=None,     # Font family
                         b=True,          # Bold (True/False)
                         i=None,          # Italics (True/False)
                         strike=None,     # strikethrough
                         outline=None,
                         shadow=None,
                         condense=None,
                         extend=None,
                         color=None,
                         u=None,
                         vertAlign=None,
                         scheme=None,
                         )
Fortunately, if you already have a Font object, you can simply initialize an InlineFont object with an existing Font object:

from openpyxl.cell.text import Font
font = Font(name='Calibri',
            size=11,
            bold=False,
            italic=False,
            vertAlign=None,
            underline='none',
            strike=False,
            color='00FF0000')
inline_font = InlineFont(font)
You can create InlineFont objects on their own, and use them later. This makes working with Rich Text cleaner and easier:

big = InlineFont(sz="30.0")
medium = InlineFont(sz="20.0")
small = InlineFont(sz="10.0")
bold = InlineFont(b=True)
b = TextBlock
rich_string2 = CellRichText(
      b(big, 'M'),
      b(medium, 'i'),
      b(small, 'x'),
      b(medium, 'e'),
      b(big, 'd')
)
For example:

red = InlineFont(color='00FF0000')
rich_string1 = CellRichText(['When the color ', TextBlock(red, 'red'), ' is used, you can expect ', TextBlock(red, 'danger')])
The CellRichText object is derived from list, and can be used as such.

Whitespace
CellRichText objects do not add whitespace between elements when rendering them as strings or saving files.

t = CellRichText()
t.append('xx')
t.append(TextBlock(red, "red"))
You can also cast it to a str to get only the text, without formatting.

str(t)
'xxred'
Editing Rich Text
As editing large blocks of text with formatting can be tricky, the as_list() method returns a list of strings to make indexing easy.

l = rich_string1.as_list()
l
['When the color ', 'red', ' is used, you can expect ', 'danger']
l.index("danger")
3
rich_string1[3].text = "fun"
str(rich_string1)
'When the color red is used, you can expect fun'
Rich Text assignment to cells
Rich Text objects can be assigned directly to cells

from openpyxl import Workbook
wb = Workbook()
ws = wb.active
ws['A1'] = rich_string1
ws['A2'] = 'Simple string'
Logo

