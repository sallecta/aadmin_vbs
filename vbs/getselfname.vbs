Dim scriptName : scriptName = WScript.ScriptName
' scriptName = Left(scriptName, Len(scriptName) - 4)
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
Dim fullpath
fullpath = fso.GetAbsolutePathName(scriptName)
WScript.Echo "fullpath:    " & fullpath & ", scriptName: " & scriptName & vbCrLf & "fullname:    " & fullpath &  "\" & scriptName
