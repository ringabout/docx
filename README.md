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

Output:

```text
长记曾携手处，千树压、西湖寒碧。
I strove with none.
For none was worth my strife;
Nature I lov’d,
And next to Nature, Art;
I warm’d both hands before the fire of life;
It sinks,
and I am ready to depart.
仰天大笑出门去，我辈岂是蓬蒿人。
```


Only parse pure text.

```nim
import docx


echo parsePureText("test.docx")
```

Output:

```text
长记曾携手处，千树压、西湖寒碧。I strove with none.For none was worth my strife;Nature I lov’d,And next to Nature, Art;I warm’d both hands before the fire of life;It sinks,and I am ready to depart.仰天大笑出门去，我辈岂是蓬蒿人。
```

Extract Picture from docx

```nim
let tmpDir = getTempDir()
if existsDir(tmpDir / "generate"):
  removeDir(tmpDir / "generate")
extractPicture("tests/test_pic.docx", tmpDir / "generate")
assert existsFile(tmpDir / "generate/image1.jpeg")
```