inputpath = "Enter your Letter|Enter your Network drive path|Enter Domani\Usaername|Enter Password"
Function test(inputpath)
on error resume next
strLetter = split(InputPath, "|")(0)
strNDPath = split(InputPath, "|")(1)
strUsername = split(InputPath, "|")(2)
strPassword = split(InputPath, "|")(3)
Set WshNetwork = WScript.CreateObject("WScript.Network")
set filesys=CreateObject("Scripting.FileSystemObject")
If filesys.FolderExists(strLetter) Then
else
WshNetwork.MapNetworkDrive strLetter, strNDPath, "FALSE", strUsername, strPassword
End If
if err.number <> 0 Then
strError = err.number
test = strError
else
test = "Success"
End if
msgbox test
End Function
call test(inputpath)