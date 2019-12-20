import os, parsexml, streams, strutils

import zip / zipFiles


let
  # UpperLetters = {'A' .. 'Z'}
  TestFile = "./test.docx"
  TempDir* = getTempDir() / "docx_windx_tmp"

template `=?=`(a, b: string): bool =
  cmpIgnoreCase(a, b) == 0

proc matchKindName(x: XmlParser, kind: XmlEventKind, name: string): bool {.inline.} =
  x.kind == kind and x.elementName =?= name

proc extractXml*(src: string, dest: string = TempDir) =
  var z: ZipArchive
  if not z.open(src):
    echo "Opening zip failed"
    quit(1)
  z.extractAll(dest)
  z.close()
  assert existsDir(dest / "word")
  assert existsFile(dest / "word/document.xml")

proc parsePureText*(fileName: string): string =
  # unpack docx
  extractXml(fileName)
  defer: removeDir(TempDir)
  let fileName = TempDir / "word/document.xml"
  # open xml file
  var s = newFileStream(fileName, fmRead)
  if s == nil: quit("cannot open the file" & fileName)
  var x: XmlParser
  defer: x.close()
  open(x, s, fileName, {reportWhitespace})

  while true:
    x.next()
    case x.kind
    of xmlCharData, xmlWhitespace:
      result &= x.charData
    of xmlEof:
      break
    else:
      discard
      


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
  open(x, s, fileName, {reportWhitespace})

  while true:
    x.next()
    if x.matchKindName(xmlElementOpen, "w:p"):
      while true:
        # ignore <w:p>
        x.next()
        if x.matchKindName(xmlElementStart, "w:t"):
          # ignore <w:t>
          x.next()
          while true:
            case x.kind
            of xmlCharData, xmlWhitespace:
              result &= x.charData
            else:
              break
            # ignore </w:t>
            x.next() 
        elif x.matchKindName(xmlElementOpen, "w:t"):
          # ignore <w:t>
          x.next()
          while x.kind != xmlElementClose:
            x.next()
          # ignore >
          x.next()
          while true:
            case x.kind
            of xmlCharData, xmlWhitespace:
              result &= x.charData
            else:
              break
            # ignore </w:t>
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
  open(x, s, fileName, {reportWhitespace})

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
          while true:
            case x.kind
            of xmlCharData, xmlWhitespace:
              res &= x.charData
            else:
              break
            # ignore </w:t>
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
  assert existsFile(TestFile)
  echo TempDir
  echo parsePureText("test.docx")
  echo parseDocument("test.docx")
  for line in docLines("test.docx"):
    echo line
