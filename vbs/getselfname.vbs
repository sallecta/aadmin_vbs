Dim scriptName : scriptName = WScript.ScriptName
' scriptName = Left(scriptName, Len(scriptName) - 4)
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
Dim fullname
fullname = fso.GetAbsolutePathName(scriptName)
Dim fullpath
fullpath = Left(fullname, Len(fullname) - Len(scriptName))
WScript.Echo "scriptName: " & scriptName & vbCrLf & "fullname:    " & fullname & vbCrLf & "fullpath: " & fullpath
