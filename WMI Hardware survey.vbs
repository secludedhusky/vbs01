	
Dim WMI, Col, Ob, S2, i2, s3, sFil, sBul, sLine

'-- path to save data.  ------------------
sFil = "C:\Sysinfo.txt"

sBul = "   " & Chr(149) & "   "
sLine = vbCrLf & "_____________________________________________" & vbCrLf & vbCrLf

Err.Clear
On Error Resume Next
  
Set WMI = GetObject("WinMgmts:")

    If (Err.number <> 0) Then 
       MsgBox "Error creating WMI object. Error: " & Err.Number & " - " & Err.Description
       WScript.quit
    End If   
      

'-------------- product ------------------------------------

Set Col = WMI.ExecQuery("Select * from Win32_ComputerSystemProduct")
  S2 = S2 & sBul & " Product Info:" & vbCrLf & vbCrLf
  For Each Ob in Col 
    S2 = S2 & "Product Name: " & Ob.Name & vbCrLf
    S2 = S2 & "Product Version: " & Ob.Version & vbCrLf
    S2 = S2 & "Product Description: " & Ob.Description & vbCrLf
    S2 = S2 & "IdentifyingNumber: " & Ob.IdentifyingNumber & vbCrLf
   S2 = S2 & "Product UUID: " & Ob.UUID & vbCrLf
  Next
S2 = S2 & sLine


 '-- box id --------------------------------------------------------------------

Set Col = WMI.ExecQuery("Select * from Win32_SystemEnclosure")
 S2 = S2 & sBul & "Machine ID (SystemEnclosure) info:" & vbCrLf & vbCrLf
For Each Ob in Col
    S2 = S2 & "Part Number: " & Ob.PartNumber & vbCrLf
    S2 = S2 & "Serial Number: " & Ob.SerialNumber & vbCrLf
     S2 = S2 & "Asset Tag: " & Ob.SMBIOSAssetTag & vbCrLf 
Next

S2 = S2 & sLine

 '-- motherboard --------------------------------------------------------------------

Set Col = WMI.ExecQuery("Select * from Win32_MotherboardDevice")
 S2 = S2 & sBul & "Motherboard info:" & vbCrLf & vbCrLf
For Each Ob in Col
    S2 = S2 & "Caption: " & Ob.Caption & vbCrLf
    S2 = S2 & "InstallDate: " & Ob.InstallDate & vbCrLf
    S2 = S2 & "DeviceID: " & Ob.DeviceID & vbclrf
Next
S2 = S2 & vbCrLf

'----------- bios -----------------------------------

Set Col = WMI.ExecQuery("Select * from Win32_BIOS")
 S2 = S2 & sBul & "BIOS info:" & vbCrLf & vbCrLf
For Each Ob in Col
    S2 = S2 & "Manufacturer: " & Ob.Manufacturer & vbCrLf
    S2 = S2 & "Description: " & Ob.Description & vbCrLf
    S2 = S2 & "Version: " & Ob.Version & vbCrLf
    S2 = S2 & "InstallDate: " & Ob.InstallDate & vbCrLf
    S2 = S2 & "SerialNumber: " & Ob.SerialNumber & vbCrLf

Next
S2 = S2 & sLine


 '-- CPU --------------------------------------------------------------------

Set Col = WMI.ExecQuery("Select * from Win32_Processor")
  S2 = S2 & sBul & "CPU:" & vbCrLf & vbCrLf
  For Each Ob in Col
    S2 = S2 & "Manufacturer: " & Ob.Manufacturer & vbCrLf
    S2 = S2 & "Description: " & Ob.Description & vbCrLf
    S2 = S2 & "Name: " & Ob.Name & vbCrLf
    S2 = S2 & "Speed: " & Ob.MaxClockSpeed & sLine
 Next

 '-- RAM and product info. --------------------------------------------------------------------

Set Col = WMI.ExecQuery("Select * from Win32_ComputerSystem")
  S2 = S2 & sBul & "Installed RAM: " 
  For Each Ob in Col 
  i2 = Ob.TotalPhysicalMemory
  If i2 > 0 Then
      i2 = i2 \ 1024 \ 1024 
      S2 = S2 & CStr(i2) & " MB" & vbCrLf
  End If    
     S2 = S2 & sLine
  Next
  
   S2 = S2 & sBul & "PC Info.:" & vbCrLf &vbCrLf

  For Each Ob in Col
   S2 = S2 & "PC or motherboard model: " & Ob.Model & vbCrLf
   S2 = S2 & "System name: " & Ob.Name & vbCrLf
   S2 = S2 & "System Manufacturer: " & Ob.Manufacturer & vbCrLf
  Next
S2 = S2 & sLine

'---------- onboard devices ----------------------------------------

Set Col = WMI.ExecQuery("Select * from Win32_OnBoardDevice")
 S2 = S2 & sBul & "Onboard devices:" & vbCrLf & vbCrLf
For Each Ob in Col
    S2 = S2 & "Description: " & Ob.Description & vbCrLf
    S2 = S2 & "Name: " & Ob.Name & vbCrLf
    S2 = S2 & "Manufacturer: " & Ob.Manufacturer & vbCrLf
    S2 = S2 & "Model: " & Ob.Model & vbCrLf & vbCrLf
 Next
S2 = S2 & sLine

'-- graphics --------------------------------------------------------------------

Set Col = WMI.ExecQuery("Select * from Win32_VideoController")
 S2 = S2 & sBul & "Graphics:" & vbCrLf & vbCrLf
For Each Ob in Col
    S2 = S2 & "Description: " & Ob.Description & vbCrLf
    S2 = S2 & "Name: " & Ob.Name & vbCrLf
    i2 = Ob.AdapterRAM
     If i2 > 0 Then
       i2 = i2 \ 1024 \ 1024
       S2 = S2 & "RAM: " & " MB" & vbCrLf
     End If
    S2 = S2 & "Driver Date: " & Ob.DriverDate & vbCrLf
    S2 = S2 & "Driver Version: " & Ob.DriverVersion & vbCrLf
Next
S2 = S2 & sLine

 '-- hard disks --------------------------------------------------------------------

Set Col = WMI.ExecQuery("Select * from Win32_DiskDrive")
  S2 = S2 & sBul & "Drives:" & vbCrLf & vbCrLf
For Each Ob in Col
        S2 = S2 & "Description: " & Ob.Description & vbCrLf
        S2 = S2 & "Manufacturer: " & Ob.Manufacturer & vbCrLf
        S2 = S2 & "Model: " & Ob.Model & vbCrLf
S2 = S2 & "InterfaceType: " & Ob.InterfaceType & vbCrLf
S2 = S2 & "MediaType: " & Ob.MediaType & vbCrLf
         S2 = S2 & "DeviceID: " & Ob.DeviceID & vbCrLf
        S2 = S2 & "Number of Win Partitions: " & Ob.Partitions & vbCrLf
        s3 = CStr(Ob.Size)
         If Len(s3) > 9 Then 
            s3 = Left(s3, (len(s3) - 9))
            S2 = S2 & "Size (GB): " & s3 
         End If
         S2 = S2 & vbCrLf & vbCrLf   
Next
S2 = S2 & sLine

 '-- CD/DVD drives --------------------------------------------------------------------

Set Col = WMI.ExecQuery("Select * from Win32_CDROMDrive")
  S2 = S2 & sBul & "CD/DVD drives:" & vbCrLf & vbCrLf
   For Each Ob in Col
     S2 = S2 & "Description: " & Ob.Description & vbCrLf
     S2 = S2 & "Caption: " & Ob.Caption & vbCrLf
     S2 = S2 & "Manufacturer: " & Ob.Manufacturer & vbCrLf & vbCrLf
  Next
S2 = S2 & sLine

'-- network adapter --------------------------------------------------------------------

Set Col = WMI.ExecQuery("Select * from Win32_NetworkAdapter")
 S2 = S2 & sBul & "Network Adapter:" & vbCrLf & vbCrLf
For Each Ob in Col
    S2 = S2 & "Description: " & Ob.Description & vbCrLf
    S2 = S2 & "Name: " & Ob.ProductName & vbCrLf
    S2 = S2 & "Manufacturer: " & Ob.Manufacturer & vbCrLf
    S2 = S2 & "MAC Address: " & Ob.MACAddress & vbCrLf & vbCrLf
 Next
S2 = S2 & sLine

Set Col = Nothing
Set WMI = Nothing

Dim FSO, TS
Set FSO = CreateObject("Scripting.FileSystemObject")
  Set TS = FSO.CreateTextFile(sFil, True)
   TS.Write S2
   TS.Close
  Set TS = Nothing
Set FSO = Nothing
 MsgBox "Done."





