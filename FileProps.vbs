'-- Drop a file onto this script to get properties. 

Dim FSO, oFil, s2, Arg
  Set FSO = CreateObject("Scripting.FileSystemObject")
   Arg = WScript.Arguments(0)
      If FSO.FileExists(Arg) = False Then
         MsgBox "Drop a file onto this script to get properties.", 64
         Set FSO = Nothing
         WScript.Quit
      End If   
  
 s2 = FSO.GetFileName(Arg)  
       MsgBox "GetFileName returns: " & s2   '-- returns file name.
       
s2 = FSO.GetBaseName(Arg)  
      MsgBox "GetBaseName returns: " & s2   '-- returns last part of path: file name without extension, or folder name.
      
s2 = FSO.GetExtensionName(Arg)
       MsgBox "GetExtensionName returns: " & s2  '--return file extension.
       
 s2 = FSO.GetFileVersion(Arg)
       MsgBox "GetFileVersion returns: " & s2  & VBCrLf & VBCrLf & "[Note: File version only applies to some file types.]"
       
 s2 = FSO.GetParentFolderName(Arg)
       MsgBox "GetParentFolderName returns: " & s2  '--return name of parent folder.
      
  
     s2 = "The File object returned by GetFile has methods to copy, move, or delete it." & VBCrLf
     s2 = s2 & "The File object also has several collections and properties. For this file the properties are as follows...."
       MsgBox s2
     
   Set oFil = FSO.GetFile(Arg)
      s2 = "Attributes: " & oFil.Attributes & VBCrLf
      s2 = s2 & "DateCreated: " & oFil.DateCreated & VBCrLf
      s2 = s2 & "DateLastAccessed: " & oFil.DateLastAccessed & VBCrLf
      s2 = s2 & "DateLastModified: " & oFil.DateLastModified & VBCrLf
      s2 = s2 & "Name: " & oFil.Name & VBCrLf
      s2 = s2 & "Size: " & oFil.Size & " bytes" & VBCrLf
      s2 = s2 & "Path: " & oFil.Path & VBCrLf
      s2 = s2 & "Type: " & oFil.Type & VBCrLf
      s2 = s2 & "ShortName: " & oFil.ShortName & VBCrLf
      s2 = s2 & "ParentFolder: " & oFil.ParentFolder & VBCrLf & VBCrLf
    
       MsgBox s2
     
    Set oFil = Nothing     
    Set FSO = Nothing
     
WScript.Quit
      
      
      