'--use this file to change the attributes of any file or folder.
'--options are normal, read-only, hidden and system.
'--any combination of attributes may be set.

Dim fso, fil, f, attr, i, arg
Set fso = CreateObject("Scripting.FileSystemObject")

  '--check for drag-and-drop:
   
     If wscript.arguments.count = 0 then
           arg = inputbox("Enter the path of the file or folder you want to change the attributes of.", "Change File Attributes")
     else
           arg = wscript.arguments.item(0)
     end if

if arg = "" then 
  wscript.quit
end if

  '--check for existence of either file or folder and keep track of which with i :

if fso.fileexists(arg) = true then 
  f = "File" 
  i = 1
elseif fso.folderexists(arg) = true then
  f = "Folder"
  i = 2
else
  msgbox "Path is wrong. No such file or folder.", 16, "Wrong Path"
  wscript.quit
end if

attr = inputbox("Enter the attributes number. Add numbers to get the desired combination of attributes." & vbcrlf & "Normal is 0" & vbcrlf & "ReadOnly is 1" & vbcrlf & "Hidden is 2" & vbcrlf & "System is 4", "Choose Attributes Setting", "0")

'--make sure that input is a number from 0 to 7:

if isnumeric(attr) = false then
    msgbox "Wrong entry. Must be a number from 0 to 7.", 16, "Wrong Number"
    wscript.quit
end if
if cint(attr) > 7 or cint(attr) < 0 then
    msgbox "Wrong entry. Only 0 to 7 are possible.", 16, "Wrong Number"
    wscript.quit
end if

'attr = cint(attr)
  if i = 1 then
     set fil = fso.GetFile(arg)
     fil.attributes = attr
     set fil = nothing
     msgbox f & " attributes changed.", 0, "All Done"
  end if
  if i = 2 then
     set fil = fso.GetFolder(arg)
     fil.attributes = attr
     set fil = nothing
    msgbox f & " attributes changed.", 0, "All Done"
 end if