Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
Dim scriptName : scriptName = WScript.ScriptName
Dim fullpath
fullpath = fso.GetAbsolutePathName("..\..")
WScript.Echo "fullpath:    " & fullpath & ", scriptName: " & scriptName
