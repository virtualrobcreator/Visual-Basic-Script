inputpath = "Enter your folder path"
Function latestfile(inputpath)
dim fileSystem, folder, file, t, list, dict, path  
path = inputpath 
Set fileSystem = CreateObject("Scripting.FileSystemObject")
Set list = CreateObject("System.Collections.ArrayList")
Set dict = CreateObject("Scripting.Dictionary")
Set folder = fileSystem.GetFolder(path) 
for each file in folder.Files
CurrentTime = dateadd("h", -0, Now) 
FileTime = file.DateLastModified
if instr(file, "CSA Salaried Headcount for RPA - Weekly Output") then        
if FileTime < CurrentTime then            
t =  CurrentTime - FileTime
list.Add t
dict.Add t, file
end if
end if
next 
list.Sort
a = list.Item(0)
latestfilename = dict(a)
latestfile = latestfilename
msgbox latestfile
END Function
call latestfile(inputpath)
'msgbox latestfile