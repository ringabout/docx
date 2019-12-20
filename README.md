# docx
A dead simple docx reader. 

Read pure text from docx written by Nim.

## Usage

Keep only newline information.

```nim
import docx


echo parseDocument("test.docx")
for line in docLines("test.docx"):
  echo line
```

Only parse pure text.

```nim
import docx


echo parsePureText("test.docx")
```