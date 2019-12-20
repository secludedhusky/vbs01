'--ClsTextStream - Textstream operations. This Class deals with reading, writing and appending.
 '--it could be expanded to add specialized text functions.
 '--This is a class "block". It can be used by simply pasting everything from "Class ClsTS" to
'-- "End Class" at the end of a script. The class can then be created as an object.

Dim CT, s, r
 Set CT = new ClsTS
   r = CT.AppendToFile("C:\windows\desktop\file1.txt", "blah blah")
 Set CT = Nothing

MsgBox r

'----------------------------------------------------
Class ClsTS

Private FSO

 Private Sub Class_Initialize()
    Set FSO = CreateObject("Scripting.FileSystemObject") 
 End Sub
          
 Private Sub Class_Terminate()
    Set FSO = Nothing
 End Sub
 
'------ Read text file. Returns file text. --------------------
'--Ex.: s = Cls.ReadFile("C:\File1.txt")
Public Function ReadFile(sFilePath)
Dim TS
  If FSO.FileExists(sFilePath) = False Then
     ReadFile = ""
     Exit Function
  End If
  
 Set TS = FSO.OpenTextFile(sFilePath, 1)
    ReadFile = TS.ReadAll
 Set TS = Nothing   

End Function
  
'------ Replace text file. sFilePath is path of file. sText is text of new file.
'-- This Sub will Write the new file and Set its attributes to whatever they
'-- were originally.
'-- ex.: Cls.Replacefile "C:\file1.txt", s

Public Sub ReplaceFile(sFilePath, sText)
Dim Atr, oFil, TS
     '--confirm valid path:
   If FSO.GetParentFolderName(sFilePath) = "" Then Exit Sub
   
      '--Set attributes to 0 and delete:
    If FSO.FileExists(sFilePath) Then
         Set oFil = FSO.GetFile(sFilePath)
            Atr = oFil.Attributes
            oFil.Attributes = 0
         Set oFil = Nothing
           FSO.DeleteFile sFilePath, True
    End If     
    
     Set TS = FSO.CreateTextFile(sFilePath, True)
       TS.Write sText
       TS.Close
     Set TS = Nothing
     
     Set oFil = FSO.GetFile(sFilePath)
        oFil.Attributes = Atr
     Set oFil = Nothing
 End Sub 
 
'---------- Append text to text file -----------------
 '-- Adds text to the End of an existing text file.
 '-- ex.: r = Cls.AppendToFile("C:\file1.txt", s)  '-- returns -1 If file does Not exist.  
 
Public Function AppendToFile(sFilePath, sText)  
Dim Atr, oFil, TS

           '--confirm file exists:
    If FSO.FileExists(sFilePath) = False Then
        AppendToFile = -1
        Exit Function
    End If
    
          '--remove Read-only If necessary:
     Set oFil = FSO.GetFile(sFilePath)
        Atr = oFil.Attributes
        oFil.Attributes = 0
    Set oFil = Nothing
    
    Set TS = FSO.OpenTextFile(sFilePath, 8)
       TS.Write sText
       TS.Close
    Set TS = Nothing
    
    Set oFil = FSO.GetFile(sFilePath)
        oFil.Attributes = Atr
    Set oFil = Nothing
  AppendToFile = 0
      
End Function

End Class