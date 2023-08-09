inputpath = "Your Folder Path"
Function latestfilename(inputpath)
Dim latestFile, latestDate, fso, folder, file, fileNameWithoutExtension, fileDate, regex, matches, fileDateStr
Set fso = CreateObject("Scripting.FileSystemObject")
Set folder = fso.GetFolder(inputpath)
For Each file In folder.Files
    ' Extract date from file name using regular expression
    ' msgbox file.name 
    If instr(file,"Your File Name Prefix")Then
        Set regex = New RegExp
        regex.Pattern = "(\d{4}-\d{2}-\d{2})"
        Set matches = regex.Execute(file.Name)
        
        If matches.Count > 0 Then
            fileDateStr = matches(0).value
            fileDate = CDate(fileDateStr)

            'msgbox fileDate 
            
            ' Compare dates to find the latest file
            If latestFile = "" Or fileDate > latestDate Then
                latestDate = fileDate
                latestFile = file.Path
            End If
        End If
    End If
Next

If latestFile <> "" Then
     latestfilename = latestFile
    'WScript.Echo "Date in the file name: " & latestDate
Else
    latestfilename = "unable to get file"
    ' WScript.Echo "No files found in the folder."
End If
msgbox latestfilename
End Function
call latestfilename(inputpath)
'msgbox latestfile