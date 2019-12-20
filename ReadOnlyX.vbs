'--Enter the path to a folder.
'--this will remove Read Only attribute from all files in folder entered.
'--the script goes down to the 4th level in folders, removing all attributes
'--in all files. (read-only, sys, archive and hidden)
'--it will also work on a single file.
'---------------------------------
Dim fso, fol, fil, r, fils, filfull, fols, fols1, fols2, fol1, fol2, fol3, arg, i
Set fso = CreateObject("Scripting.FileSystemObject")

 If wscript.arguments.count = 0 then
           arg = inputbox("Enter folder path for Read-Only removal.", "READ-ONLY REMOVAL")
     else
           arg = wscript.arguments.item(0)
     end if

if arg = "" then 
wscript.quit
end if
if fso.fileexists(arg) = true then 
 i = 1
elseif fso.folderexists(arg) = true then
i = 2
else
 msgbox "Path is wrong. No such file or folder.", 16, "Wrong Path"
 wscript.quit
end if
'--if it's a file.: ----------------------------
if i = 1 then
   set fol = fso.getfile(arg)
   fol.attributes = 0
  set fol = nothing
  wscript.quit
end if
'--if it's a folder: --------------------------
set fol = fso.GetFolder(arg)
set fils = fol.files
on error resume next
 folprocess  '-- do first level

set fols = fol.Subfolders    '--do 2nd level
 if fols.count = 0 then
  msgbox "All Done", 0
  wscript.quit
 end if
  for each fol1 in fols
    set fol = fso.GetFolder(fol1)
    folprocess
         set fols1 = fol.subfolders  '--do 3rd level
          if fols1.count <> 0 then
            for each fol2 in fols1
              set fol = fso.GetFolder(fol2)
              folprocess
                 set fols2 = fol.subfolders  '--do 4th level
                    if fols2.count <> 0 then
                       for each fol3 in fols2
                          set fol = fso.GetFolder(fol3)
                          folprocess
                          set fol = nothing
                       next
                    end if
              set fol = nothing
            next
       end if
  next
msgbox "All done.", 0
wscript.quit

sub folprocess()
set fils = fol.files
for each fil in fils              '-- do files in folder
  set filfull = fso.GetFile(fil)
  filfull.attributes = 0
 set filfull = nothing
next
set fils = nothing
end sub