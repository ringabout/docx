import unittest

import docx


suite  "test parse text":
  test "can parse pure text":
    echo parsePureText("tests/test.docx")

  test "can parse":
    echo parseDocument("tests/test.docx")

  test "can read line":
    for line in docLines("tests/test.docx"):
      echo line
