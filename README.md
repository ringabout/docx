# docx
A dead simple docx reader. 

Read pure text from docx written by Nim.Keep only newline information.

## Usage
```nim
import docx


echo parseDocument("test.docx")
for line in docLines("test.docx"):
  echo line
```