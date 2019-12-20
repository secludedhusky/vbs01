'MsgBox GetCUAppDataPath()

Dim AppDataFol, FPath, IEUPath, FSO, Pt1, LocFol, WASFil, sApFol, sMsg, Ret, FoxPath
Dim AllUsersAppDataFol, LocalAppDataFol, DocsFol, AllUsersDocsFol, BooFlash
 Err.Clear
 On Error Resume Next

Ret = MsgBox("Be sure to read the Storage Data Info file before running this script. Do you want to proceed now?", 292)
   If Ret = 7 Then WScript.Quit
   
AppDataFol = GetFolderPathWSlash(26)  'app data

AllUsersAppDataFol = GetFolderPathWSlash(35)  'all users app data.

LocalAppDataFol = GetFolderPathWSlash(28)  'local app data.

DocsFol = GetFolderPathWSlash(5)  'docs.

AllUsersDocsFol = GetFolderPathWSlash(46) ' all users docs.
  

Set FSO = CreateObject("Scripting.FileSystemObject")
        
'----------------- Adobe Flash ---------------------------------

BooFlash = False
DelFlash AppDataFol 'flash cookies normally here. Other folder paths added for good measure.
DelFlash AllUsersAppDataFol
DelFlash LocalAppDataFol
' DelFlash DocsFol  ' commented because some people could have a folder that they've named "Macromedia" in their documents folder.
DelFlash AllUsersDocsFol
 
If BooFlash = False Then MsgBox "No Flash cookies found.", 64


 '----------------------- IE path -------------------------------------
     
Err.Clear
 On Error Resume Next

IEUPath = AppDataFol & "Microsoft\Internet Explorer\UserData"
   FSO.DeleteFolder IEUPath, True
 
  Select Case Err.Number
   Case 0
     sMsg = "Successfully deleted hidden Internet Explorer data."
   Case 76, 53
     sMsg = "No hidden Internet Explorer data folder found."
   Case Else
     sMsg = "Unable to delete hidden Internet Explorer data folder. Reason: "
     sMsg = sMsg & vbCrLf & Err.Description
End Select       
 
MsgBox sMsg, 64

'---------------- firefox --------------------------
  
 '-- 2 possible locations for Firefox file. Check both.

Err.Clear
On Error Resume Next
Dim OFols, oFol, oFol1, oFils, oFil, FolName
FoxPath = AppDataFol & "Mozilla\Firefox\Profiles"
WASFil = "\webappsstore.sqlite"
    Set oFol = FSO.GetFolder(FoxPath)
      Set OFols = oFol.SubFolders
         For Each oFol1 in OFols
             FolName = oFol1.Name
                If FSO.FileExists(FoxPath & "\" & FolName & WASFil) = True Then
                    FSO.DeleteFile FoxPath & "\" & FolName & WASFil, True
                    If Err.number = 0 Then
                       MsgBox "Successfully deleted hidden Firefox data.", 64
                    Else   
                      sMsg = "Unable to delete hidden Firefox data folder. Reason: "
                      sMsg = sMsg & vbCrLf & Err.Description
                      MsgBox sMsg, 64
                    End If
                  DropIt  '-- if the firefox file was found then quit.
                End If
           Next   
                   
   FoxPath = LocalAppDataFol & "Mozilla\Firefox\Profiles"
    Set oFol = FSO.GetFolder(FoxPath)
      Set OFols = oFol.SubFolders
         For Each oFol1 in OFols
             FolName = oFol1.Name
                If FSO.FileExists(FoxPath & "\" & FolName & WASFil) = True Then
                    FSO.DeleteFile FoxPath & "\" & FolName & WASFil, True
                    If Err.number = 0 Then
                       MsgBox "Successfully deleted hidden Firefox data.", 64
                    Else   
                      sMsg = "Unable to delete hidden Firefox data folder. Reason: "
                      sMsg = sMsg & vbCrLf & Err.Description
                      MsgBox sMsg, 64
                    End If
                  DropIt  '-- if the firefox file was found then quit.
                End If
           Next   

  '-- If still going at this point then no webappsstore.sqlite file was found.

   MsgBox "No hidden Firefox data found.", 64
  DropIt

Sub DropIt()
  Set oFils = Nothing
  Set oFol1 = Nothing
  Set OFols = Nothing
  Set oFol = Nothing
  Set FSO = Nothing
  WScript.Quit
End Sub

'-------------------------------- function to return system folder path. ------------
 Function GetFolderPathWSlash(iFol)
   Dim sPathC, ShApC, FolOb
          On Error Resume Next
       Set ShApC = CreateObject("Shell.Application")
       Set FolOb = ShApC.NameSpace(iFol)  '--Shell.Application Namespace method to get folder path.
         GetFolderPathWSlash = FolOb.self.path
         If Right(GetFolderPathWSlash, 1) <> "\" Then GetFolderPathWSlash = GetFolderPathWSlash & "\"
       Set FolOb = Nothing
       Set ShApC = Nothing
  End Function
  
' check for/delete flash cookies.
Sub DelFlash(sPath)
 On Error Resume Next
  FSO.DeleteFolder sPath & "Macromedia", True
Select Case Err.Number
   Case 0
     sMsg = "Successfully deleted hidden Adobe Flash data from:" & vbCrLf & sPath
     BooFlash = True
     MsgBox sMsg
   Case 76, 53
       'sMsg = "No hidden Adobe Flash data folder found."
   Case Else
     sMsg = "Unable to delete hidden Adobe Flash data folder from:" &  vbCrLf & sPath & vbCrLf & "Reason: "
     sMsg = sMsg &  Err.Description
     MsgBox sMsg
End Select       
End Sub