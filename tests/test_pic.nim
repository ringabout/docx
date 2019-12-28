import unittest, os

import docx



suite "test extract picture":
  test "can extract picture":
    let tmpDir = getTempDir()
    if existsDir(tmpDir / "generate"):
      removeDir(tmpDir / "generate")
    extractPicture("tests/test_pic.docx", tmpDir / "generate")
    check existsFile(tmpDir / "generate/image1.jpeg")


