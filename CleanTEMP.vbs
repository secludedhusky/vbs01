
'-- Script to clean out TEMP files.
'-- This script searches for possible TEMP folder locations on all Windows versions:
'--  Paths checked are like this (specific path may vary by ssystem):  
'    C:\%WIN%\TEMP
'    C:\TEMP
'    C:\Documents and Settings\[folder name]\Local Settings\TEMP
'       In the Documents and Settings folder the script looks through 
'       each subfolder for a TEMP folder. You must have full permission
'       on NT systems for this to work.
'
'--  Cleaning NT systems is really the whole point of this script. On Win9x
'   there is usually just one TEMP folder, in the Windows folder. But on NT
'   (NT4, 2000, XP) there may be C:\WINNT\TEMP and C:\TEMP, as well as
'   a TEMP folder for each user! When the API is used to find the TEMP folder
'   it returns only the current user's TEMP folder. Moreover, there is no API
'   to return an array or collection of all possible TEMP folders. So the only
'   way to properly clean leftover TEMP files is to run a script like this while
'   logged on with full permission.

Dim sList, FSO2, sFol, Pt1, Pt2, Pt3, sCU, sBase, sTempPath, sPath1, oFol4, oFols4, OFolSub, Ret

On Error Resume Next

Ret = MsgBox("This script will delete files from all TEMP folders it can find. Other software should be closed first. Also, on NT systems (2000, XP) the script should be run by an administrator with full access to the file system." & vbCrLf & vbCrLf & "Do you want to proceed?", 33, "TEMP Cleaner")
  If Ret = 2 Then WScript.Quit
  
'-- check for %WIN%\TEMP

Set FSO2 = CreateObject("Scripting.FileSystemObject")
sFol = FSO2.GetSpecialFolder(0) & "\TEMP"
 If (FSO2.FolderExists(sFol) = True) Then
     sList = CleanFiles(sFol) & vbCrLf
     sPath1 = sFol
 End If

'-- check home drive ( ex.: C:\TEMP )
sFol = Left(sFol, 3) & "temp"
  If FSO2.FolderExists(sFol) = True Then
     sPath1 = sFol
     sList = sList & CleanFiles(sFol) & vbCrLf
  End If
  
'-- Get TEMP folder as the system sees it. On Win9x this is typically C:\Windows\TEMP.
'-- On WinNT systems, unfortunately, the TEMP path is the current users TEMP folder
'-- and there is no API for returning all TEMP folders.

sFol =  FSO2.GetSpecialFolder(2)
  If (FSO2.FolderExists(sFol) = True) And (sFol <> sPath1) Then
     sList = sList & CleanFiles(sFol) & vbCrLf
  End If
  
'-- Find other user's, and all users, TEMP folders by walking the Docunemts and Settings folder tree.
'-- There seems to be no official way to find the Documents and Settings folder, so this script is
'-- presuming it's a parent folder of the current user's TEMP folder.

Pt1 = InStr(1, sFol, ":\Docum", 1)
Pt2 = InStr((Pt1 + 3), sFol, "\")
Pt3 = InStr((Pt2 + 1), sFol, "\")

'-- If no more paths found then quit here. It's probably either Win9x or someone running
'-- without permission.

  If (Pt1 = 0) Or  (Pt2 = 0) Or  (Pt3 = 0) Then
    sList = "TEMP folders found: List shows beginning size of each TEMP folder found and size of that folder after cleaning." & vbCrLf & vbCrLf & sList
     MsgBox sList
     Set FSO2 = Nothing
     WScript.Quit
  End If

'-- go through documents and settings subfolders to find TEMP
'-- folder for each user.

sBase = Left(sFol, Pt2)  ' c:\documents and settings\
sCU = Left(sFol, Pt3 - 1)
sCU = Right(sCU, (len(sCU) - Len(sBase)))

Set oFol4 = FSO2.GetFolder(sBase)
  Set oFols4 = oFol4.subfolders
     For Each OFolSub in oFols4
        If (OFolSub.Name <> sCU) Then
           sFol = sBase & OFolSub.name & "\Local Settings\Temp"
           If FSO2.FolderExists(sFol) = True Then
              sList = sList & CleanFiles(sFol) & vbCrLf
           End If
        End If
     Next
  Set oFols4 = Nothing
Set oFol4 = Nothing   

   sList = "TEMP folders found: List shows beginning size of each TEMP folder found and size of that folder after cleaning." & vbCrLf & vbCrLf & sList
   MsgBox sList
   Set FSO2 = Nothing
   WScript.Quit


'-- END OF SCRIPT ------------------------------------------

'-- Below here is the function called to delete files in each found TEMP folder.
'-- This function is fairly simple. It just deletes subfolders and files, then formats
'-- a return string that reports folder size before deletions and folder size after deletions.

Function Cleanfiles(Path)
   Dim FSO, oFol, oFol2, oFols, oFils, oFil, Sz1, Sz2, Szi1, Szi2
      Set FSO = CreateObject("Scripting.FileSystemObject")
 On Error Resume Next
Set oFol = FSO.GetFolder(Path)
Sz1 = oFol.Size
  Set oFols = oFol.SubFolders
    For Each oFol2 in oFols
      oFol2.Delete True
    Next
  Set oFols = Nothing   
  Set oFils = oFol.Files
    For Each oFil in oFils
      oFil.Delete True
    Next
  Set oFils = Nothing
 Sz2 = oFol.Size
Set oFol = Nothing     
Set FSO = Nothing

Szi1 = " Bytes"
If (Sz1 > 1024) Then
  Sz1 = Sz1 \ 1024
 Szi1 = " KB"
End If
If (Sz1 > 1024) Then
  Sz1 = Sz1 \ 1024
  Szi1 = " MB"
End If

Szi2 = " Bytes"
If (Sz2 > 1024) Then
  Sz2 = Sz2 \ 1024
  Szi2 = " KB"
End If
If (Sz2 > 1024) Then
  Sz2 = Sz2 \ 1024
  Szi2 = " MB"
End If

CleanFiles = Path & ": " & Sz1 & Szi1 & " - " & Sz2 & Szi2 & vbCrLf

End Function