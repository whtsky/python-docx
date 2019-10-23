# python-docx-whtsky

![https://travis-ci.com/whtsky/python-docx.svg?branch=master](https://travis-ci.com/whtsky/python-docx)

_python-docx-whtsky_ is a Python library for creating and updating Microsoft Word
(.docx) files.

More information is available in the [python-docx documentation](https://python-docx.readthedocs.org/en/latest/)

## Release History

### 0.8.10.2 (2019-10-23)

- Add ability to restart list numbering. ( https://github.com/python-openxml/python-docx/pull/210 )

<details>
<summary>Example</summary>

https://github.com/python-openxml/python-docx/issues/25#issuecomment-143231954

```python
from docx import Document

document = Document()

# Add desired numbering styles to your template file.

# Extract abstractNumId from there. In this example, abstractNumId is 10

numId = document.get_new_list("10")

# Add a list

p = document.add_paragraph(style = 'ListParagraph', text = "a")
p.num_id = numId
p.level = 0
p = document.add_paragraph(style = 'ListParagraph', text = "b")
p.num_id = numId
p.level = 1
p = document.add_paragraph(style = 'ListParagraph', text = "c")
p.num_id = numId
p.level = 1
p = document.add_paragraph(style = 'ListParagraph', text = "d")
p.num_id = numId
p.level = 0
p = document.add_paragraph(style = 'ListParagraph', text = "e")
p.num_id = numId
p.level = 1
p = document.add_paragraph(style = 'ListParagraph', text = "f")
p.num_id = numId
p.level = 0

# Restart numbering at the outer level

numId = document.get_new_list("10")

# Add the same list once again. The numbering is restarted

p = document.add_paragraph(style = 'ListParagraph', text = "a")
p.num_id = numId
p.level = 0
p = document.add_paragraph(style = 'ListParagraph', text = "b")
p.num_id = numId
p.level = 1
p = document.add_paragraph(style = 'ListParagraph', text = "c")
p.num_id = numId
p.level = 1
p = document.add_paragraph(style = 'ListParagraph', text = "d")
p.num_id = numId
p.level = 0
p = document.add_paragraph(style = 'ListParagraph', text = "e")
p.num_id = numId
p.level = 1
p = document.add_paragraph(style = 'ListParagraph', text = "f")
p.num_id = numId
p.level = 0

document.save("num.docx")

```

</details>

### 0.8.10.1 (2019-10-16)

- allow table looking (header row/col, footer row/col, bands) modification. https://github.com/python-openxml/python-docx/pull/579
- Added font property to paragraph. https://github.com/python-openxml/python-docx/pull/417
