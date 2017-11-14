Dim scriptName : scriptName = WScript.ScriptName
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
Dim fullname
fullname = fso.GetAbsolutePathName(scriptName)
Dim fullnameExe
fullnameExe = Left(fullname, Len(fullname) - 4) & ".exe"
WScript.Echo "fullnameExe:    " & fullnameExe
sParams = ""
with CreateObject("WScript.Shell")
  .Run fullnameExe & " " & sParams, 1, true ' Wait for finish or False to not wait
end with
' wsh.echo "Done"
