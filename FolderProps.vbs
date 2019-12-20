'-- Folder object. Drop a folder on this script to get properties.
 '-- Dropped folder should have subfolders and files inside.
 
Dim FSO, s2, oFol, oFol2, oFols, oFil, oFils, Arg

  Set FSO = CreateObject("Scripting.FileSystemObject")
     Arg = WScript.Arguments(0)
       If FSO.FolderExists(Arg) = False Then
          MsgBox "Drop a folder on this script to get folder properties.", 64
          Set FSO = Nothing
          WScript.Quit
       End If
       
     
      s2 = FSO.GetParentFolderName(Arg)  
      MsgBox "GetParentFolderName returns: " & s2   '-- returns last part of path: file name without extension, or folder name.
      
   s2 = "The Folder object returned by GetFolder has several properties." & VBCrLf 
   s2 = s2 & "The properties for this folder are as follows..."
       MsgBox s2
       
  Set oFol = FSO.GetFolder(Arg)
    
      s2 = "Attributes: " & oFol.Attributes & VBCrLf
      s2 = s2 & "DateCreated: " & oFol.DateCreated & VBCrLf
      s2 = s2 & "DateLastAccessed: " & oFol.DateLastAccessed & VBCrLf
      s2 = s2 & "DateLastModified: " & oFol.DateLastModified & VBCrLf
      s2 = s2 & "Name: " & oFol.Name & VBCrLf
      s2 = s2 & "Size: " & oFol.Size & " bytes" & VBCrLf
      s2 = s2 & "Path: " & oFol.Path & VBCrLf
      s2 = s2 & "Type: " & oFol.Type & VBCrLf
      s2 = s2 & "ShortName: " & oFol.ShortName & VBCrLf
      s2 = s2 & "ParentFolder: " & oFol.ParentFolder & VBCrLf
      s2 = s2 & "Is a root folder: " & oFol.IsRootFolder & VBCrLf & VBCrLf
       MsgBox s2
     
     s2 = "The Folder object also has a Subfolders collection and a Files collection..."
     MsgBox s2
          
s2 = ""
   Set oFols = oFol.SubFolders
      For Each oFol2 in oFols    '-- enumerate subfolders using for/each. Each oFol2 returned is a folder object.
         s2 = s2 & oFol2.Name & VBCrLf
      Next
    Set oFols = Nothing
       s2 = "These are the names of the subfolders in " & oFol.Path & ":" & VBCrLf & VBCrLf & s2
       MsgBox s2
       
s2 = ""
   Set oFils = oFol.Files
      For Each oFil in oFils  '-- enumerate files in the folder using For/Each. Each oFil is a File object.
         s2 = s2 & oFil.Name & VBCrLf
      Next   
        s2 = "These are the names of the files in " & oFol.Path & ":" & VBCrLf & VBCrLf & s2
       MsgBox s2
   Set oFils = Nothing    
   
    Set oFol = Nothing
    Set FSO = Nothing  
WScript.Quit
      
      
      