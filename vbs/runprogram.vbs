Dim scriptName : scriptName = WScript.ScriptName
scriptName = Left(scriptName, Len(scriptName) - 4)
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
Dim fullnameExe
fullnameExe = fso.GetAbsolutePathName("..\..") & "\" & scriptName & ".exe"
' WScript.Echo "fullnameExe:    " & fullnameExe
sParams = ""
with CreateObject("WScript.Shell")
  .Run fullnameExe & " " & sParams, 1, true ' Wait for finish or False to not wait
end with
' wsh.echo "Done"
