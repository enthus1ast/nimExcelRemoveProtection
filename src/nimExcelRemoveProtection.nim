import zippy, cligen
import strutils, os
import std/prelude
import zippy/ziparchives

iterator allSheets(): string =
  for path in walkPattern(getAppDir() / "unzipped" / "xl" / "worksheets" / "sheet*.xml"):
    yield path

proc breakProtection(sheet, data: string): string =
  const starter = "<sheetProtection"
  const ender = "/>"
  result = data
  var pos = data.find starter
  if pos > -1:
    var endPos = data.find(ender, pos + starter.len)
    if endPos == -1:
      echo "Ender not found... error"
      return
    endPos.inc ender.len
    result.delete(pos..endPos - 1)

proc remove(path: string) =
  if not path.endswith(".xlsx"):
    echo "not an excel file (.xlsx missing)"
    return
  if not path.fileExists():
    echo "File does not exists: ", path
    return
  try:
    extractAll(path, getAppDir() / "unzipped")
  except:
    echo "Could not extract excel file: ", path
    return
  for sheet in allSheets():
    var data = readFile(sheet)
    let dataBroken = breakProtection(sheet, data)
    writeFile(sheet, dataBroken)
  let fileParts = splitFile(path)
  let newPath = fileParts.dir / fileParts.name & "_broken" & fileParts.ext
  let unzipPath = getAppDir() / "unzipped/"
  createZipArchive(unzipPath, newPath)
  echo "Written to: ", newPath
  removeDir(unzipPath)

when isMainModule:
  dispatch(remove)

