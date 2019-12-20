'-- Script to convert self-executing CAB file (EXE) to CAB file.
'-- Just drop the EXE file onto script to produce a CAB file.

'-- A S.E. CAB file is a CAB with an EXE "stub" prepended. The
'-- executed and it unpacks the CAB, in addition to doing any other 
'-- operations written into the stub. The beginning of the CAB file
'-- starts with 8 distinctive bytes: 77(M) - 83(S) - 67(C) - 70(F) - 0 - 0 - 0 - 0
'-- The M-S-C-F bytes typically appear in the stub, but not followed
'-- by 4 zero bytes. So this script works by reading the CAB to find
'-- the CAB file marker, then writing only that part and the following bytes 
'-- to disk, effectively snipping off the EXE stub.

'-- NOTE: Generally a self-executing CAB is from Microsoft, and their
'-- stub seems to always be about 60-67 KB bytes. To make the script
'-- as fast as possible, it starts by checking for the file marker between
'-- 58000 and 70000 bytes. If not found, the search is widened to
'-- 30000 to 100000 bytes. If the file marker is still not found then the function
'-- will return error #4. In that case it may be that the file is:
'-- 1) not a self-executing CAB, 2) is a self-executing CAB but has an exotic, 
'-- large stub.

Dim cc, Arg, Ret
  
   Arg = WScript.Arguments(0)
     If Len(Arg) = 0 Then
       MsgBox "Drop a self-executing CAB file [EXE] onto script for conversion.", 64, "EXE2CAB"
       WScript.Quit
     End If
     
    Set cc = new ClsCAB
      Ret = cc.EXE2CAB(Arg)
      If Ret = 0 Then
         Arg = Left(Arg, (len(Arg) - 3)) & "cab"
         MsgBox "CAB file saved as " & Arg, 64, "EXE2CAB"
      Else
         MsgBox "Error number: " & CStr(Ret)
      End If
    Set cc = Nothing  
          
Class ClsCAB
  Dim cFSO, Char1, cTS

Public Function EXE2CAB(sEXEPath)
  Dim LFil, s2, sTag, LPt, sFil, sCABPath, LStart, LEnd
    EXE2CAB = 1  '-- invalid path.
  If cFSO.FileExists(sEXEPath) = False Then Exit Function
    EXE2CAB = 2  '-- not an EXE.
  If UCase(Right(sEXEPath, 4)) <> ".EXE" Then Exit Function 
  
  LFil = GetSize(sEXEPath)
    EXE2CAB = 3 '-- file too small.   
  If LFil < 100000 Then Exit Function
  
   EXE2CAB = 4   '-- CAB file marker not found.
   
  LStart = 58000
  LEnd = 70000
  sTag = "MSCF" & Char1 & Char1 & Char1 & Char1

   sFil = Read(sEXEPath, LStart, LEnd)
   s2 = GetByteString(sFil)
   LPt = InStr(s2, sTag)
   
     If LPt = 0 Then 
        LStart = 30000
        LEnd = 100000
        sFil = Read(sEXEPath, LStart, LEnd)
        s2 = GetByteString(sFil)
        LPt = InStr(s2, sTag)
     End If

      If LPt = 0 Then Exit Function
    
  LPt = LStart + LPt '-- offset of actual CAB file start.
  sCABPath = Left(sEXEPath, (len(sEXEPath) - 3)) & "cab"
  sFil = Read(sEXEPath, (LPt - 1), (LFil - LPt) + 2)

      Set cTS = cFSO.CreateTextFile(sCABPath, True)
         cTS.Write sFil
         cTS.Close
      Set cTS = Nothing       
 
 EXE2CAB = 0     
End Function

Private Function Read(sFilePath, StartPt, LenR)
      Dim LenF
   On Error Resume Next
     Read = ""
       If (cFSO.FileExists(sFilePath) = False) Then Exit Function
     LenF = GetSize(sFilePath)
         If (StartPt >= LenF) Then Exit Function                                                            
      If (StartPt < 1) Then StartPt = 1   
      If (LenR = 0) Then LenR = LenF      
     Set cTS = cFSO.OpenTextFile(sFilePath, 1)            
         If (StartPt > 1) Then cTS.Skip (StartPt - 1)
         Read = cTS.Read(LenR)
         cTS.Close
     Set cTS = Nothing                                                       
End Function

  '-- substitute 1 for chr(0) to make string readable.
Private Function GetByteString(sStr)
  Dim sRet, iLen, iA, iLen2, A2()
     
     ReDim A2(Len(sStr) - 1)
        For iLen = 1 to Len(sStr) 
           iA = Asc(Mid(sStr, iLen, 1))
              If iA = 0 Then 
                 A2(iLen - 1) = Char1
              Else
                 A2(iLen - 1) = Chr(iA)
              End If
        Next         
            GetByteString = Join(A2, "")
End Function

Private Function GetSize(sFilePath)
  Dim cOFil
   If (cFSO.FileExists(sFilePath) = False) Then
       GetSize = -1
       Exit Function
   End If
     Set cOFil = cFSO.GetFile(sFilePath)
       GetSize = cOFil.Size
     Set cOFil = Nothing
End Function

Private Sub Class_Initialize()
       sAst = "*"
       Char1 = Chr(1)
      Set cFSO = CreateObject("Scripting.FileSystemObject")
  End Sub
          
  Private Sub Class_Terminate()
      Set TS = Nothing   '-- just in case.
      Set cFSO = Nothing
  End Sub
  
End Class
