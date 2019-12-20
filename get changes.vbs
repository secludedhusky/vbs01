'--enter a date formatted as 2/3/00, 10/21/00, etc.
'--a file will be created on the desktop named changes.txt that lists all system files
'--with a Date Created or Date Last Modified that matches that date.

Dim fso, fol, fils, fil, r, dc, dlm, fname, t, s, fil2, sh, deskfol
Set fso = CreateObject("Scripting.FileSystemObject")
Set sh = CreateObject("WScript.Shell") 
  '--get a date. (If Cancel is clicked it will cause an error.)
On Error Resume Next
    r = inputbox("Enter date to check for new files.", "Sys. Changes")

   '--create a file to write matches and get the System folder.
  
      deskfol = sh.SpecialFolders("Desktop")
      set t = fso.CreateTextFile(deskfol & "\changes.txt")
      set fol = fso.GetSpecialFolder(1)
      set fils = fol.files

on error resume next
  for each fil2 in fils  '--for each file in system folder files collection....

      fil = fso.GetAbsolutePathName(fil2)
                     '--get date created and strip off time:
           dc = fil2.DateCreated
             s = InStr(1, dc, " ", 1)
             dc = left(dc, s - 1)
                   '--get date last modified and strip off time:
          dlm = fil2.datelastmodified
             s = InStr(1, dlm, " ", 1)
            dlm = left(dlm, s - 1)

         '--if either date matches input, write filename and dates to file:
   
         if r = dc or r = dlm then      
            fname = fso.GetFilename(fil2)
            t.write fname & vbcrlf & "created " & dc & VbCrLf & "modified " & dlm & VbCrLf & VbCrLf
        end if
 
  next
 
t.close
 set t = nothing
set fol = nothing
set fils = nthing
set fso = nothing
msgbox "All done.", 0
wscript.quit
  



