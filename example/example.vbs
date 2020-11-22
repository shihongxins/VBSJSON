Sub Import(filename)
    Dim WshShell : Set WshShell = WScript.CreateObject("WScript.Shell")
    Dim objFs : Set objFs = CreateObject("Scripting.FileSystemObject")
    filename = WshShell.ExpandEnvironmentStrings(filename)
    filename = objFs.GetAbsolutePathName(filename)
    strCode = objFs.OpenTextFile(filename).ReadAll
    ExecuteGlobal strCode
End Sub

Import "../release/vbsJSON.vbs"

Dim FSO
Set FSO = CreateObject("Scripting.FileSystemObject")
Dim source1, source2
source1 = FSO.OpenTextFile("source1.json").ReadAll
source2 = FSO.OpenTextFile("source2.json").ReadAll

Dim JSON
Set JSON = new vbsJSON

Dim source1_parse, source2_parse
Set source1_parse = JSON.parse(source1)
source2_parse = JSON.parse(source2)

Dim source1_parse_stringify, source2_parse_stringify
source1_parse_stringify = JSON.stringify(source1_parse)
source2_parse_stringify = JSON.stringify(source2_parse)