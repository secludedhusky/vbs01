'-- This is a barebones Base64 encoder/decoder. Drop a file onto script and click YES
'-- to encode. Click NO to decode a Base64 string.
'-- This script uses only VBS and FileSystemObject to do its work. The basic function
' of Base64 conversion is to take each 3 bytes of binary data and convert it to 4
' 6-bit units, which allows any data to be stored as plain text because on plain
' text ASCII characters are used. Decoding is the reverse.
' FSO is designed to only handle text data. Special treatment is required to handle
' binary data, but FSO *can* do it. For example, Textstream.ReadAll expects to read
' a string, so it will return file bytes up until the first null byte. But Textstream.Read(length-of-file)
' can be used to read in the entire file as a string, regardless the content. The bytes can
' then be handled by using Asc to convert the string into a numeric array. It's inefficient,
' but it works. When the file is written back to disk the array members are then converted
' back to characters and the whole thing is transferred as a string. That works fine as
' long as one doesn't try to handle it as a string. For instance, checking Len of the string
' returned from DecodeBase64 will only return the position of the first null.
' The vbCrLf option with encoding is to accomodate email, which by tradition 
' inserts a return every 76 characters. In other words, these functions can be used
' to create or decode attachments in email. They could also be used to send any type
' of file in the form of text pasted into an email. If the recipient has the decode script
' they can just select and copy the email content, paste it into Notepad, save it as a
' TXT file, then drop it onto the script to convert that text into the original JPG, EXE, or 
' any other file type.

Dim FSO, TS, sIn, sOut, Arg, IfEncode, OFil, LSize, LRet

Arg = WScript.Arguments(0)

LRet = MsgBox("Click yes to encode file or no to decode.", 36)
  If LRet = 6 Then 
      IfEncode = True
  Else
      IfEncode = False
  End If    

Set FSO = CreateObject("Scripting.FileSystemObject")
Set OFil = FSO.GetFile(Arg)
LSize = OFil.Size
Set OFil = Nothing
Set TS = FSO.OpenTextFile(Arg)
sIn = TS.Read(LSize)
Set TS = Nothing

If IfEncode = True Then
    sOut = ConvertToBase64(sIn, True)
     Set TS = FSO.CreateTextFile(Arg & "-64", True)
         TS.Write sOut
         TS.Close
      Set TS = Nothing 
Else
    sOut = DecodeBase64(sIn)
     Set TS = FSO.CreateTextFile(Arg & "-de64", True)
         TS.Write sOut
         TS.Close
      Set TS = Nothing 
End If

Set FSO = Nothing

MsgBox "Done."
'------------------------------------------------------
Function ConvertToBase64(sBytes, AddReturns)
  Dim B2(), B76(), ABytes(), ANums
  Dim i1, i2, i3, LenA, NumReturns, sRet
     On Error Resume Next
      ANums = Array(65, 66, 67, 68, 69, 70, 71, 72, 73, 74, 75, 76, 77, 78, 79, 80, 81, 82, 83, 84, 85, 86, 87, 88, 89, 90, 97, 98, 99, 100, 101, 102, 103, 104, 105, 106, 107, 108, 109, 110, 111, 112, 113, 114, 115, 116, 117, 118, 119, 120, 121, 122, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 43, 47)
  
     LenA = Len(sBytes)
       '-- convert each string character to ASCII value.
     ReDim ABytes(LenA - 1)
       For i1 = 1 to LenA
           ABytes(i1 - 1) = Asc(Mid(sBytes, i1, 1))
       Next  
      '-- generate base 64 equivalent in array B2.
  ReDim Preserve ABytes(((LenA - 1) \ 3) * 3 + 2) 
  ReDim Preserve B2((UBound(ABytes) \ 3) * 4 + 3) 
     i2 = 0
        For i1 = 0 To (UBound(ABytes) - 1) Step 3
            B2(i2) = ANums(ABytes(i1) \ 4)
              i2 = i2 + 1
            B2(i2) = ANums((ABytes(i1 + 1) \ 16) Or (ABytes(i1) And 3) * 16)
              i2 = i2 + 1
            B2(i2) = ANums((ABytes(i1 + 2) \ 64) Or (ABytes(i1 + 1) And 15) * 4)
              i2 = i2 + 1
            B2(i2) = ANums(ABytes(i1 + 2) And 63)
              i2 = i2 + 1
        Next 
            For i1 = 1 To i1 - LenA
               B2(UBound(B2) - i1 + 1) = 61 ' add = signs at end if necessary.
            Next 
            
      '-- Most email programs use a maximum of 76 characters per line when encoding
      '-- binary files as base 64. This next function achieves that by generating another
      '--- array big enough for the added vbCrLfs, then copying the base 64 array over.
      
   If (AddReturns = True) And (LenA > 76) Then
        NumReturns = ((UBound(B2) + 1) \ 76)
        LenA = (UBound(B2) + (NumReturns * 2)) '--make B76 B2 plus 2 spots for each vbcrlf.
         ReDim B76(LenA)
          i2 = 0
          i3 = 0
              For i1 = 0 To UBound(B2)
                   B76(i2) = B2(i1)
                    i2 = i2 + 1
                    i3 = i3 + 1
                       If (i3 = 76) And (i2 < (LenA - 2)) Then   '--extra check. make sure there are still
                          B76(i2) = 13                 '-- 2 spots left for return if at end.
                          B76(i2 + 1) = 10
                          i2 = i2 + 2
                          i3 = 0
                       End If
              Next
        For i1 = 0 to UBound(B76)
            B76(i1) = Chr(B76(i1))
        Next        
          sRet = Join(B76, "")
   Else
        For i1 = 0 to UBound(B2)
            B2(i1) = Chr(B2(i1))
        Next  
          sRet = Join(B2, "")
   End If
       ConvertToBase64 = sRet
End Function

Function DecodeBase64(Str64)
  Dim B1(), B2()
  Dim i1, i2, i3, LLen, UNum, s2, sRet, ANums
  Dim A255(255)
    On Error Resume Next
        ANums = Array(65, 66, 67, 68, 69, 70, 71, 72, 73, 74, 75, 76, 77, 78, 79, 80, 81, 82, 83, 84, 85, 86, 87, 88, 89, 90, 97, 98, 99, 100, 101, 102, 103, 104, 105, 106, 107, 108, 109, 110, 111, 112, 113, 114, 115, 116, 117, 118, 119, 120, 121, 122, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 43, 47)
   
    For i1 = 0 To 255
       A255(i1) = 64
    Next
    For i1 = 0 To 63
       A255(ANums(i1)) = i1
    Next
          s2 = Replace(Str64, vbCr, "")
          s2 = Replace(s2, vbLf, "")
          s2 = Replace(s2, " ", "")
          s2 = Trim(s2)
          LLen = Len(s2)
         ReDim B1(LLen - 1)
      For i1 = 1 to LLen
          B1(i1 - 1) = Asc(Mid(s2, i1, 1)) 
      Next      

  '--B1 is now in-string as array.
   ReDim B2((LLen \ 4) * 3 - 1)
        i2 = 0
     For i1 = 0 To UBound(B1) Step 4
        B2(i2) = (A255(B1(i1)) * 4) Or (A255(B1(i1 + 1)) \ 16)
           i2 = i2 + 1
        B2(i2) = (A255(B1(i1 + 1)) And 15) * 16 Or (A255(B1(i1 + 2)) \ 4)
           i2 = i2 + 1
        B2(i2) = (A255(B1(i1 + 2)) And 3) * 64 Or A255(B1(i1 + 3))
           i2 = i2 + 1
     Next
        If B1(LLen - 2) = 61 Then
           i2 = 2
        ElseIf B1(LLen - 1) = 61 Then
           i2 = 1
        Else
           i2 = 0
        End If
        UNum = UBound(B2) - i2
     ReDim Preserve B2(UNum)
       For i1 = 0 to UBound(B2)
         B2(i1) = Chr(B2(i1))
       Next   
        DecodeBase64 = Join(B2, "")
End Function
