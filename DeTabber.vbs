'-- This script will remove horizontal and vertical tabs from a text file.
'-- it then saves the result as a new file, in the same folder, with the same
'-- name except that "DeT" is added. ex.: file.txt becomes DeTfile.txt
'-- To use it, if you have at least v. 5.1 of the in. Scripting Host you can
'-- just "drop" the file on this script. Otherwise, double-click the script
'-- and enter the path to the file.
'--------------------------------------------------------------
Dim fso, ts, arg, folpath, s, newpath, newfil, ext
Set fso = CreateObject("Scripting.FileSystemObject")

     If wscript.arguments.count = 0 then
           arg = inputbox("Enter path to file for detabbing.", "Detabber", "C:\Windows\Desktop\")
     else
           arg = wscript.arguments.item(0)
     end if
  if fso.FileExists(arg) = false then
     msgbox "Wrong path.", 64
     wscript.Quit
  end if
 
  set ts = fso.OpenTextFile(arg, 1, false)
    s = ts.ReadAll
    ts.Close
  set ts = nothing

s = Replace(s, chr(9), "", 1, -1, 0)
s = Replace(s, chr(11), "", 1, -1, 0)

  
  folpath = fso.GetParentFolderName(arg)
   newfil = fso.GetFileName(arg)
  newfil = folpath & "\DeT" & newfil
  
 set ts = fso.CreateTextFile(newfil, true)
  ts.Write s
  ts.Close
 set ts = nothing
 

