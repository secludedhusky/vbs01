'--Enter date formatted as 2/3/00, 12/23/00, etc. 
'--enter folder path.
'--this script will go down 4 levels to check date created and date last modified.
'--any file found that has a matching date for either property will be recorded in a file,
'--changes.txt, on the desktop. 


Dim FSO, fol, s, r, r2, t, filfull, fols, fols1, fols2, fol1, fol2, fol3, deskpath
Set FSO = CreateObject("Scripting.FileSystemObject")
Dim SH
Set SH = CreateObject("WScript.Shell")
r = InputBox("Enter date for file changes.", "File Changes")
r2 = InputBox("Enter folder to look in.", "File Changes")
   If FSO.FolderExists(r2) = False Then
       MsgBox "Wrong path", 0
       WScript.quit
   End If
deskpath = SH.SpecialFolders("Desktop")
deskpath = deskpath & "\changes.txt" 

Set t = FSO.CreateTextFile(deskpath, True)

Set fol = FSO.GetFolder(r2)
'--set fils = fol.files
On Error Resume Next
 folprocess fol '-- do first level

Set fols = fol.Subfolders    '--do 2nd level
 If fols.count = 0 Then
   MsgBox "All Done", 0
   t.Close
   Set t = Nothing
   WScript.quit
 End If
  For Each fol1 in fols
    Set fol = FSO.GetFolder(fol1)
    folprocess fol
         Set fols1 = fol.subfolders  '--do 3rd level
          If fols1.count <> 0 Then
            For Each fol2 in fols1
              Set fol = FSO.GetFolder(fol2)
              folprocess fol
                 Set fols2 = fol.subfolders  '--do 4th level
                    If fols2.count <> 0 Then
                       For Each fol3 in fols2
                          Set fol = FSO.GetFolder(fol3)
                          folprocess fol
                          Set fol = Nothing
                       Next
                    End If
              Set fol = Nothing
            Next
       End If
  Next
   t.Close
   Set t = Nothing
MsgBox "All done.", 0
WScript.quit

Sub folprocess(obfol)
Dim fil, fil1, dc, dlm, pfil, fils
 Set fils = obfol.files
 For Each fil in fils              '-- do files in folder
    pfil = FSO.GetAbsolutePathName(fil)
     Set fil1 = FSO.GetFile(pfil)
     dc = fil1.DateCreated
       s = InStr(1, dc, " ", 1)
       dc = left(dc, s - 1)
    dlm = fil1.datelastmodified
       s = InStr(1, dlm, " ", 1)
       dlm = left(dlm, s - 1)
         If r = dc or r = dlm Then 
            t.write pfil & VBCrLf & "created " & dc & VBCrLf & "modified " & dlm & VBCrLf & VBCrLf
        End If
    Set fil1 = Nothing
  Next
   Set fils = Nothing
End Sub