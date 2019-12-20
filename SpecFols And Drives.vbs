'-- FileSystemObject demo.
 '-- shows how to get paths to system folders.
Dim FSO, s2, oDrv, oDrvs

  Set FSO = CreateObject("Scripting.FileSystemObject")
      s2 = "Windows Folder: " & FSO.GetSpecialFolder(0) & VBCrLf  '-- Windows folder on Win9x. WINNT folder on WinNT.
      s2 = s2 & "System Folder: " & FSO.GetSpecialFolder(1) & VBCrLf   '-- System folder on Win9x. System32 folder on WinNT.
      s2 = s2 & "TEMP Folder: " & FSO.GetSpecialFolder(2) & VBCrLf
       
    MsgBox "These are the special folder paths:" & VBCrLf & s2
    
    MsgBox "Next is a demonstration of the Drives collection and Drive object.", 64
    On Error Resume Next
  Set oDrvs = FSO.Drives  '-- drives collection
    For Each oDrv in oDrvs  '-- enumerate drives. each one is a Drive object.
         If oDrv.IsReady Then
              s2 = "DriveLetter: " & oDrv.DriveLetter & VBCrLf
              s2 = s2 & "AvailableSpace: " & oDrv.AvailableSpace & VBCrLf
              s2 = s2 & "DriveType: " & oDrv.DriveType & VBCrLf
              s2 = s2 & "FreeSpace: " & oDrv.FreeSpace & VBCrLf
              s2 = s2 & "TotalSize: " & oDrv.TotalSize & VBCrLf
              s2 = s2 & "VolumeName: " & oDrv.VolumeName & VBCrLf & VBCrLf
              
               MsgBox s2
         End If      
     Next
   
    Set oDrvs = Nothing
    Set FSO = Nothing  
WScript.Quit
      
      
      