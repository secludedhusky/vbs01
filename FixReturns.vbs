''------ FixReturns.vbs can be used to correct text files that lack carriage returns, such
'--  as text from a Unix-created file.
'------ These files will have normal text but show boxes where there should be a return.
'-- this script will automatically process either a dropped file or all files in a dropped folder.
'--- BE CAREFUL NOT TO DROP BINARY FILES THAT MIGHT POSSIBLY GET CORRUPTED BY THIS PROCESS.

Dim FSO, TS, s, Arg, fil, FPath, s1, oFol, oFils, oFil, Ret
Set FSO = CreateObject("Scripting.FileSystemObject")
     If WScript.arguments.count = 0 Then
           Arg = InputBox("This script will correct web server text that lacks carriage returns. Enter path of file.", "Fix File", "C:\Windows\Desktop\")
     Else
           Arg = WScript.arguments.item(0)
     End If
     
     If FSO.FolderExists(Arg) = True Then
       Ret = MsgBox("All files in folder must be plain text. Binary files could be corrupted by this function. Do you still want to proceed?", 33, "Caution:")
           If Ret = 2 Then
               Set FSO = Nothing
               WScript.Quit
           End If
           
          Set oFol = FSO.GetFolder(Arg)
            Set oFils = oFol.Files
               For Each oFil in oFils
                   FPath = oFil.Path
                   Set oFil = Nothing
                   FixReturnsFile FPath
               Next
            Set oFils = Nothing
          Set oFol = Nothing
          
     ElseIf FSO.FileExists(Arg) = True Then
         FixReturnsFile Arg
     Else
         MsgBox "Wrong path.", 64, "No such file"       
     End If
  
Set FSO = Nothing
WScript.quit

Sub FixReturnsFile(sPath)
  On Error Resume Next
   Set TS = FSO.OpenTextFile(sPath, 1, False)
       s = TS.ReadAll
       TS.Close
   Set TS = Nothing

 '-------- replace linefeed characters with vbcrlf ------------------------
s1 = Replace(s, vbCrLf, vbCr, 1, -1, 0)  '-- reoplace all vbcrlf with vbCr.
s1 = Replace(s1, vbLf, vbCr, 1, -1, 0)  '-- now any vbLf are alone, so also replace those.
s1 = Replace(s1, vbCr, vbCrLf, 1, -1, 0)  '-- now any vbCr should represent a single lie return, whether Unix, MS Word, etc.
 
'-- -----write file. -----------------
If FSO.fileexists(sPath) = True Then
  FSO.deletefile sPath, True
End If
  Set TS = FSO.CreateTextFile(sPath, True)
     TS.Write s1
     TS.Close
  Set TS = Nothing
End Sub  



