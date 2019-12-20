'-- Script to convert simple DOC files to text.

Dim Arg, TFile, CBin, s2, A1, AText(), A2(1), A4(3), i2, i3, LStart, LEnd, LLen, CR2, CrLf2, Ret, sErrStr, UB, Boo1, BooTog
  
 On Error Resume Next

   CR2 = vbCr & vbCr
  CrLf2 = vbCrLf & vbCrLf
  
  Arg = WScript.Arguments(0)
    If Len(Arg) = 0 Then
       MsgBox "This script works by dropping a .DOC file onto it.", 64, "DOCtoTXT"
       WScript.Quit
    End If
    
    If UCase(Right(Arg, 4)) <> ".DOC" Then
       MsgBox "This script works by dropping a .DOC file onto it.", 64, "DOCtoTXT"
       WScript.Quit
    End If

    TFile = Left(Arg, (len(Arg) - 3)) & "txt" '-- path for text file.
    

    Set CBin = New ClsBin
    s2 = CBin.Read(Arg, 513, 32)
    
     If (len(s2) = 0) Then
       MsgBox "Error. File seems to be invalid.", 64, "DOCtoTXT"
       Set CBin = Nothing
       WScript.Quit
    End If

    A1 = CBin.GetArray(s2, False)
    
    If (A1(0) <> 236) Or  (A1(1) <> 165) Then
       MsgBox "Error. DOC file header FIB section not found. This does not seem to be a valid DOC file. Script cannot continue.", 64, "DOCtoTXT"
       Set CBin = Nothing
       WScript.Quit
    End If

'-- Check fComplex flag.

  A2(0) = A1(10)
  A2(1) = A1(11)
    i2 = CBin.GetNumFromBytes(A2)
   If (i2 And 4) = 4 Then
       MsgBox "This is a complex type DOC file. Script cannot continue.", 64, "DOCtoTXT"
       Set CBin = Nothing
       WScript.Quit
    End If
    
'-- check for Macintosh character set here?  '--todo
    
'-- This could stand to have more error checking here. As it is, the code
'-- assumes that once the FIB marker is found the values for text offset, text length, etc.
'-- will be valid.

'-- get text start and offset.
  A4(0) = A1(24)
  A4(1) = A1(25)
  A4(2) = A1(26)
  A4(3) = 0  '-- Text should not be over 65 KB from start of file, so just skip this.
    LStart = CBin.GetNumFromBytes(A4)
    
  A4(0) = A1(28)
  A4(1) = A1(29)
  A4(2) = A1(30)
  A4(3) = A1(31)
    LEnd = CBin.GetNumFromBytes(A4)

    LLen = (LEnd - LStart) '-- starting offset is LStart. LLen is bytes to read.
    s2 = CBin.Read(Arg, (LStart + 513), LLen) '-- add 1 to LStart because CBin is 1-based. Also add 512 for FIB offset.

    A1 = CBin.GetArray(s2, False)
    UB = UBound(A1)
    ReDim AText(UB + 1000) '-- won't need all this. just padding to be on the safe side.
 '-- s2 is now text of file. Fix it up and write to file.  
  i3 = 0
  Boo1 = True  '-- boolean to track whether to write file data.
  BooTog = False '-- used to filter double vbCr.
  
'-- In addition to a number of characters that need to be dropped or changed,
'-- "fields" need to be removed. This next section uses a tokenizing routine
'-- to walk the text string byte by byte. It's more work than a series of replace
'-- functions, but it's more flexible. With tokenizing the fields can just be
'-- dropped from the final text of the file.

    For i2 = 0 to UB
       If (Boo1 = False)  Then  
             '-- Boo1 is set to false when Chr(19) is encountered, which means the start of a field.
             '-- 21 marks end of field... 21 marks start of field text. Resume adding text to file.
             '-- So this bit here is designed to skip the character if it's in a field but toggle back to
             '-- read for the next character once Chr(21) or 20 is found.
           If (A1(i2) = 21) Or (A1(i2) = 20) Then Boo1 = True  
       Else     
            Select Case A1(i2)
                Case 19
                   Boo1 = False '-- marks beginning of field. drop out following text.
                Case 11, 12
                    AText(i3) = 13
                    AText(i3 + 1) = 10
                    i3 = i3 + 2
                Case 13
                   If BooTog = True Then
                      BooTog = False  '-- skip 2nd vbCr in series.
                   Else   '-- Chr(13) is end of paragraph, so convert it to vbCrLf & vbCrLf.
                           '-- but Chr(13) also often comes in pairs, which ends up leaving too much space.
                            '-- so convert any Chr(13) found but skip the 2nd if there are two together.
                      AText(i3) = 13
                      AText(i3 + 1) = 10
                      AText(i3 + 2) = 13
                      AText(i3 + 3) = 10
                      i3 = i3 + 4
                      BooTog = True
                   End If    
                Case 0, 1, 2, 3, 4, 5, 6, 7, 8, 14, 20, 21  
                  '-- nothing. drop these out.
                Case 30, 31, 45
                    AText(i3) = 45
                    i3 = i3 + 1
                Case 145   ' apostrophe fix where Word puts a single opening quote mark.
                    AText(i3) = 39
                    i3 = i3 + 1
                Case Else
                   AText(i3) = A1(i2)
                   i3 = i3 + 1
            End Select
        End If
          '-- un-toggle 2nd vbCr trap if this was not a vbCr
         If (A1(i2) <> 13) Then BooTog = False
    Next
    
    ReDim Preserve AText(i3 - 1)
        
Ret = CBin.WriteFileA(TFile, AText, True)
  If (Ret <> 0) Then
     MsgBox "Error " & CStr(Ret) & ". " & sErrStr, 64, "DOCtoTXT"
  Else
     MsgBox "Text version saved as" & vbCrLf & TFile & ".", 64, "DOCtoTXT"
  End If
  
 Set CBin = Nothing
 WScript.Quit


'-- This is a mini-version of ClsBin for handling binary files with Textstream.
'-- The full version is is available at www.jsware.net/jsware/scripts.php3

'---- ClsBin Class FUNCTIONS: ------------------------------------
''
'       GetSize(FilePath)  - returns size of file in bytes.
'
'       Read(FilePath, Start, Length) - returns a string from file.
'
'       GetArray(StringIn, SnipUnicode) - convert a string to an array of byte values. If SnipUnicode = True then get only every 2nd byte.
'
'      GetNumFromBytes(array) - takes array of ubound 1 or 3. return numeric value for 2 or 4 bytes.
'
'      WriteFileA(sFilePath, ArrayIn, OverWrite) - Write file with string.
'
'---------------------------------------------------------------------------

Class ClsBin
   Private FSO, i, TS, sAst, ANums, Char1
   
  Private Sub Class_Initialize()
       sAst = "*"
       Char1 = Chr(1)
      Set FSO = CreateObject("Scripting.FileSystemObject")
  End Sub
          
  Private Sub Class_Terminate()
      Set TS = Nothing   '-- just in case.
      Set FSO = Nothing
  End Sub
 
  '----------------- size = GetSize(FilePath) ---------------------------------------------------------
    '--                       get size of file in bytes. returns -1 if file not found.
Public Function GetSize(sFilePath)
  Dim OFil
   If (FSO.FileExists(sFilePath) = False) Then
       GetSize = -1
       Exit Function
   End If
     Set OFil = FSO.GetFile(sFilePath)
       GetSize = OFil.Size
     Set OFil = Nothing
End Function
 
'-- This is just a wrapper for TexStream.Read function, to simplify things 
'-- and avoid needing to deal with TS and FSO details repeatedly.
'-- note that ReadAll does not return a usable string. This function always uses Textstream.Read.
'--------------------------------- s = Read(FilePath, StartPoint, Length) ------------------------------------
Public Function Read(sFilePath, StartPt, LenR)
Dim LenF
   On Error Resume Next
     Read = ""
       If (FSO.FileExists(sFilePath) = False) Then Exit Function
     LenF = GetSize(sFilePath)
       If (StartPt >= LenF) Then Exit Function   '-- if startpoint is beyond end of file then quit.
                                                            '-- if request is to Read beyond end of file then just Read to end and return that.
      If (StartPt < 1) Then StartPt = 1       '-- adjust in case 0 was sent for start point.
      If (LenR = 0) Then LenR = LenF          '-- send 0 in 3rd parameter to Read entire file.
     Set TS = FSO.OpenTextFile(sFilePath, 1)            
         If (StartPt > 1) Then TS.Skip (StartPt - 1)
         Read = TS.Read(LenR)
         TS.Close
     Set TS = Nothing                                                       
End Function

'---------------- Write a file. -------------------------------------------
Public Function WriteFileA(sFilePath, ArrayIn, OverWrite)
 Dim sA1, iA1
     On Error Resume Next
       If (FSO.FileExists(sFilePath) = True) Then
          If (OverWrite = True) Then
              FSO.DeleteFile sFilePath, True
           Else
              WriteFileA = 1  '-- file exists.
              Exit Function
           End If  
       End If
       If IsArray(ArrayIn) = False Then
           WriteFileA = 2    '-- ArrayIn value is not an array.
           Exit Function
      End If  
   Err.Clear   
    For iA1 = 0 to UBound(ArrayIn)
        ArrayIn(iA1) = Chr(ArrayIn(iA1))
    Next    
       sA1 = Join(ArrayIn, "")
       
      Set TS = FSO.CreateTextFile(sFilePath, True)
         TS.Write sA1
         TS.Close
      Set TS = Nothing
                                   '-- return 0 if no errors.
    WriteFileA = Err.Number 
    If (Err.number <> 0) Then sErrStr = Err.Description
End Function

 '-- returns an array of byte values from a string. This is a way to leave the 0-bytes alone
 '-- while still being able to Read numeric values from the bytes.
Function GetArray(sStr, SnipUnicode)
Dim iA, Len1, Len2, AStr()
  On Error Resume Next
  Len1 = Len(sStr)
   If (SnipUnicode = True) Then 
      ReDim AStr((Len1 \ 2) - 1)
   Else
     ReDim AStr(Len1 - 1)
   End If      
 
   If (SnipUnicode = True) Then 
          For iA = 1 to Len1 step 2
             AStr(iA - 1) = Asc(Mid(sStr, iA, 1))
         Next    
   Else
         For iA = 1 to Len1
             AStr(iA - 1) = Asc(Mid(sStr, iA, 1))
         Next      
   End If  
      GetArray = AStr    
End Function
'-------------------- return a number from 2 or 4 bytes. ---------------
Public Function GetNumFromBytes(AIn)
   Dim Num1
       On Error Resume Next
   Select Case UBound(AIn)
      Case 1
        Num1 = AIn(0) + (AIn(1) * 256)
      Case 3  
        Num1 = AIn(0) + (AIn(1) * 256)
        Num1 = Num1 + (AIn(2) * 65536)
        Num1 = Num1 + (AIn(3) * 16777216)
      Case Else
        Num1 = 0
   End Select     
     If (Err.number = 0) Then
         GetNumFromBytes = Num1
     Else
         GetNumFromBytes = -1
     End If
End Function
 
End Class
    
