import os, parsexml, streams, strutils

import zip / zipFiles


let
  TempDir* = getTempDir() / "docx_windx_tmp"

template `=?=`(a, b: string): bool =
  cmpIgnoreCase(a, b) == 0

proc matchKindName(x: XmlParser, kind: XmlEventKind, name: string): bool {.inline.} =
  x.kind == kind and x.elementName =?= name
  
when defined(windows):
  {.passl: "-lz".}

proc extractXml*(src: string, dest: string = TempDir) {.inline.} =
  if not existsFile(src):
    raise newException(IOError, "No such file: " & src)
  var z: ZipArchive
  if not z.open(src):
    raise newException(IOError, "[ZIP] Can't open file: " & src)
  z.extractAll(dest)
  z.close()

proc parsePureText*(fileName: string): string {.inline.} =
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


proc parseDocument*(fileName: string): string {.inline.} =
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

proc extractPicture*(src, dest: string) =
  # unpack docx
  extractXml(src)
  defer: removeDir(TempDir)
  let fileName = TempDir / "word/media"
  if not existsDir(dest):
    createDir(dest)
  if existsDir(fileName):
    moveDir(fileName, dest)


when isMainModule:
  echo parsePureText("../../tests/test.docx")
  echo parseDocument("../../tests/test.docx")
  for line in docLines("../../tests/test.docx"):
    echo line
