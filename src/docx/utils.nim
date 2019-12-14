import os, parsexml, streams, strutils

import zip / zipFiles


let
  # UpperLetters = {'A' .. 'Z'}
  TestFile = "./test.docx"  
  TempDir = getTempDir() / "docx_windx_tmp" 
assert existsFile(TestFile)

template `=?=`(a, b: string): bool =
  cmpIgnoreCase(a, b) == 0

proc matchKindName(x: XmlParser, kind: XmlEventKind, name: string): bool {.inline.} =
  x.kind == kind and x.elementName =?= name


proc extractXml*(fileName: string) =
  var z: ZipArchive
  if not z.open(fileName):
    echo "Opening zip failed"
    quit(1)
  z.extractAll(TempDir)
  z.close()
  assert existsDir(TempDir / "word")
  assert existsFile(TempDir / "word/document.xml")

proc parseDocument*(fileName: string): string =
  # unpack docx
  extractXml(fileName)
  defer: removeDir(TempDir)
  let fileName = TempDir / "word/document.xml"
  # open xml file
  var s = newFileStream(fileName, fmRead)
  if s == nil: quit("cannot open the file" & fileName)
  var x: XmlParser
  defer: x.close()
  open(x, s, fileName)

  while true:
    x.next()
    if x.matchKindName(xmlElementOpen, "w:p"):
      while true:
        # ignore <w:p>
        x.next()
        if x.matchKindName(xmlElementStart, "w:t"):
          # ignore <w:t>
          x.next()
          while x.kind == xmlCharData:
            result &= x.charData
            x.next()
        elif x.matchKindName(xmlElementOpen, "w:t"):
          # ignore <w:t>
          x.next()
          while x.kind != xmlElementClose:
            x.next()
          # ignore >
          x.next()
          while x.kind == xmlCharData:
            result &= x.charData
            x.next()
        elif x.matchKindName(xmlElementEnd, "w:p"):
          break
        else:
          discard
      result &= "\n"
    elif x.kind == xmlEof:
      break
    else:
      discard

iterator docLines*(fileName: string): string = 
  # unpack docx
  extractXml(fileName)
  defer: removeDir(TempDir)
  let fileName = TempDir / "word/document.xml"
  # open xml file
  var s = newFileStream(fileName, fmRead)
  if s == nil: quit("cannot open the file" & fileName)
  var x: XmlParser
  defer: x.close()
  open(x, s, fileName)

  var res: string
  while true:
    x.next()
    if x.matchKindName(xmlElementOpen, "w:p"):
      res = ""
      while true:
        # ignore <w:p>
        x.next()
        if x.matchKindName(xmlElementStart, "w:t"):
          # ignore <w:t>
          x.next()
          while x.kind == xmlCharData:
            res &= x.charData
            x.next()
        elif x.matchKindName(xmlElementOpen, "w:t"):
          # ignore <w:t>
          x.next()
          while x.kind != xmlElementClose:
            x.next()
          # ignore >
          x.next()
          while x.kind == xmlCharData:
            res &= x.charData
            x.next()
        elif x.matchKindName(xmlElementEnd, "w:p"):
          break
        else:
          discard
      yield res
    elif x.kind == xmlEof:
      break
    else:
      discard

when isMainModule:
  echo TempDir
  echo parseDocument("test.docx")
  for line in docLines("test.docx"):
    echo line