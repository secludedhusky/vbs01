'-- This script measures folder sizes and then writes a report. It works recursively.
'-- That is, it will measure all folders under a given path. To measure all folders in a drive just 
'-- enter C:\, D:\, etc. when the script starts.
'-- This script converts all sizes to MB and then sorts them. The final report lists
'-- folders in order of size, with the larget first. The main point of this script is to
'-- provide the ability to do a very quick check of folders on a drive to see where
'-- unexpected bloat might be when cleaning up a system.

'-- NOTE: For accurate results, make sure that "hidden" folders are visible. On some systems
'--  hidden folders may not be seen by the script.

Dim FSO, s2, TS, sDriv, APath(), ASize(), iCnt, iTotal, i2

sDriv = InputBox("Enter path of folder to list sizes of all subfolders. For a drive enter X:\, where X is the drive letter.", "List Folder Sizes")

If Len(sDriv) < 3 Then WScript.quit

Set FSO = CreateObject("Scripting.FileSystemObject")
  
  On Error Resume Next

iTotal = 500
ReDim APath(iTotal)
ReDim ASize(iTotal)
iCnt = 0

GetFolderSizes(sDriv)

iCnt = iCnt - 1
ReDim Preserve ASize(iCnt)
ReDim Preserve APath(iCnt)

QuickSort ASize, APath, 0, 0

For i2 = iCnt to 0 step -1
      s2 = s2 & APath(i2) & " -- " & CStr(ASize(i2)) & " MB" & vbCrLf
Next   

s2 = "Sizes of folders in " & sDriv & vbCrLf & "(Note: Sizes are in MB. A size of 0 indicates the size is between 0 and 1 MB.)" & vbCrLf & vbCrLf & s2


Set TS = FSO.CreateTextFile("C:\Folder Sizes.txt", True)
   TS.Write s2
   TS.Close
Set TS = Nothing

MsgBox "List for " & sDriv & " is saved to C:\Folder Sizes.txt", 64

Sub GetFolderSizes(sPath)
   Dim oFol, oFols, oFol1, iSz, sList, s
     On Error Resume Next
     
     APath(iCnt) = sPath
      Set oFol = FSO.GetFolder(sPath)
     iSz = oFol.Size
     If iSz > 1024 Then iSz = iSz / 1024: Else iSz = 0
     If iSz > 1024 Then iSz = iSz / 1024: Else iSz = 0
      iSz = CInt(iSz) 
     ASize(iCnt) = iSz
     
      iCnt = iCnt + 1
     If iCnt + 10 > iTotal Then
         iTotal = iTotal + 200
         ReDim Preserve APath(iTotal)
         ReDim Preserve ASize(iTotal)
     End If
         
     Set oFols = oFol.SubFolders
          If oFols.count > 0 Then
             For Each oFol1 in oFols
                 GetFolderSizes(oFol1.Path)
            Next
          End If  
        Set oFols = Nothing
    Set oFol = Nothing   
 End Sub

Sub QuickSort(AIn1, AIn2, LBeg, LEnd)
  Dim LBeg2, vMid, LEnd2, vSwap1, vSwap2
    On Error Resume Next
      If (LEnd = 0) Then LEnd = UBound(AIn1)
    LBeg2 = LBeg
    LEnd2 = LEnd
     vMid = AIn1((LBeg + LEnd) \ 2)
      Do
          Do While AIn1(LBeg2) < vMid And LBeg2 < LEnd   
             LBeg2 = LBeg2 + 1
          Loop
          Do While vMid < AIn1(LEnd2) And LEnd2 > LBeg   
             LEnd2 = LEnd2 - 1
          Loop
            If LBeg2 <= LEnd2 Then
               vSwap1 = AIn1(LBeg2)
               vSwap2 = AIn2(LBeg2)

               AIn1(LBeg2) = AIn1(LEnd2)
               AIn2(LBeg2) = AIn2(LEnd2)
             
               AIn1(LEnd2) = vSwap1
               AIn2(LEnd2) = vSwap2

               LBeg2 = LBeg2 + 1
               LEnd2 = LEnd2 - 1
            End If
     Loop Until LBeg2 > LEnd2
       If LBeg < LEnd2 Then QuickSort AIn1, AIn2, LBeg, LEnd2
       If LBeg2 < LEnd Then QuickSort AIn1, AIn2, LBeg2, LEnd
End Sub 

