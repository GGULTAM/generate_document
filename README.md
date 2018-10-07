# generate_document
This is a tool generating a PDF document from a DOCX document.

# Usage
How to import
```
from generate_document import replace_docx
```

How to use
```
replace_docx(src, dest, replace_list)
```
> src is path of the template DOCX file\n

> dest is path of the output PDF file\n

> replace_list is the list of texts you want to change... like below.

```
[['original text 1', 'new text 1'], ['original text 2', 'new textg 2'], ...]
```

# Dependency
- python-docx
- MS Office Word on Windows
- pywin32 (for module 'wincom32')
