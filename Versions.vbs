'--this will create a file on the desktop named versions.txt that lists the version
'--numbers for all DLL and OCX files in System folder.

Dim fso, ext, fol, fils, fil, ts, fname, v, f, sh, deskfol
Set fso = CreateObject("Scripting.FileSystemObject")
Set sh = CreateObject("WScript.Shell")

   '--get the system folder. If a variable is used it gets the path,
   '--but if an object is used it gets the folder.
set fol = fso.GetSpecialFolder(1) 
set fils = fol.files

   '--get the desktop folder and create file:
deskfol = sh.SpecialFolders("Desktop")
set ts = fso.CreateTextFile(deskfol & "\versions.txt", True)
on error resume next
 
  '--go through system files collection. Get full path, from that get extension.
  '--if extension is dll or ocx, get the file version and rite info to file.

for each f in fils
   fil = fso.GetAbsolutePathName(f)
   ext = fso.GetExtensionName(fil)
      if ucase(ext) = "DLL" or ucase(ext) = "OCX" then
           v = fso.GetFileVersion(fil)
           fname = fso.GetFilename(fil)
           ts.writeline fname & "     " & v
     end if
next
ts.close
msgbox "All done.", 0
set ts = nothing
set fol = nothing
set fils = nothing












