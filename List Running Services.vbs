 Dim WMI, AAll(), i6, AllServs, oServ, sServ, FSO, TS

 Set WMI = GetObject("WinMgmts:")

 ReDim AAll(200)
    i6 = 0
    Set AllServs = WMI.ExecQuery("select * from Win32_Service where Started = True")  
     For Each oServ in AllServs
         sServ = oServ.Name & vbCrLf & oServ.DisplayName & vbCrLf & oServ.Description & vbCrLf & oServ.StartMode & vbCrLf & oServ.PathName & vbCrLf & "__________________________________" & vbCrLf
         AAll(i6) = sServ
         i6 = i6 + 1
     Next     
     
    ReDim Preserve AAll(i6)
    sServ = Join(AAll, vbCrLf)

  Set AllServs = Nothing    
  
Set FSO = CreateObject("Scripting.FileSystemObject")
Set TS = FSO.CreateTextFile("C:\Services Currently Running.txt", True)
  TS.Write sServ
  TS.Close
Set TS = Nothing
Set FSO = Nothing

MsgBox "Done. List of running services saved as C:\Services Currently Running.txt"

  
