' MSI Utility Class. This Class can be used to research and unpack MSI and MSM files.
 '_________________________________________________

'-- THIS FILE, MSIUnPack.vbs, IS A STAND_ALONE VERSION INTENDED FOR LINUX USERS.
'-- REQUIREMENTS: MSIUnPack.vbs and jcabxmsi.exe must be in the same folder. The Windows
'-- Script Host must be installed through WINE. (Note that some of the DLLs and OCXs in
'-- the WSH install may need to be hand-registered at the command line. That was required with Suse 10
'-- using WINE .9.10. It may have been fixed since then. (Note that WINE .9.10 comes AFTER .9.9!
'-- The WINE people don't mean "point nine, point one oh". They mean "point nine, dot ten".)

'-- Class Public Functions  -------------------------------------------------------------------------

'                  Major Functions:

 '     Boolean = Cls.ProcessMSI(filepath) - reads database of msi file and loads info into Class. also extracts any embedded CAB files.
  
 '     Boolean = Cls.UnpackMSI() - called only after ProcessMSI returns True. unpacks msi file and writes info. log.
  
 '     Boolean = Cls.WriteMSIDesc() - called only after ProcessMSI returns True. writes info. log without unpacking msi file.
  
'                  Minor Functions:

'      Boolean = Cls.GetFeatureInfo(featureName, AVar) - takes feature name and variable. If return is True Then AVar is array(1). 
          ' AVar(0) = text description of feature.  AVar(1) = bullet-divided list of components required For feature.
 
'      Boolean = Cls.GetFeatureList(AVar) - If True returns array of feature names.
         
'      Boolean = Cls.GetComponentInfo(CompID, AVar) - If True Then AVar is array(3).
             ' AVar(0) = component folder Path.  AVar(1) = bullet-divided file Path list. AVar(2) = bullet-divided file ID list.
             '   AVar(3) = VBCrLf-divided reg.setting list With bullet-divided parts.
             ' by using split on aVar(1) and aVar(2) matching arrays can be produced to Get real file names from MSI IDs.
     
'      Boolean = Cls.GetComponentList(AVar) - If True returns array of comp. names.
   
'      array(1) = FindFileComp(sFileName) - return component and feature associated With file as array(1). (0) - comp. (1) - feature.
                '     If no database loaded return "null".  If file Not matched, return " ".  If feature Not found, return " ".
                
'       Boo = GetFileList(Array)
  
'       string = GetFeatureParent(FeatureID) - return feature parent ID or "none".
'       string = GetCompCondition(CompID) - return condition For installation or "none".
                                                    
'       string = GetSummaryInfo() - return summary desc. of package.

'       string = GetCompRegList(sComp) - -return reg. list For component, separated by VBCrLf

'       string = GetAllRegList() - return all reg. settings in VBCrLf list.

'       string = TranslateRegStrings(sRegIn) - '--process registry data into shell.RegWrite-ready lines. send in VBCrLf-delimited string, Get back same.

'       Sub WriteData path, string  -  same simple Sub in both classes For convenience. That way ComClass does Not
                                             ' need to be instantiated unless it's actually being used to Get com reg data.
                                             
'       Sub AppendData path, string  - Sub to add to file. used For writing COM registry sets in Case it's too much For one string.
'
'
'

Dim ParFol, Boo1, LPt, MC, Arg
  ParFol = WScript.ScriptFullName
  LPt = InStrRev(ParFol, "\")
    ParFol = Left(ParFol, LPt) '-- get parent folder path of this running script.

  Arg = WScript.Arguments(0)
   If Len(Arg) = 0 Then
      MsgBox "Drop an MSI file onto script for unpacking.", 64, "MSI Unpacker"
      WScript.Quit
    End If
    
  Set MC = new MSIOps
    MC.CurrentFolderPath = ParFol
    Boo1 = MC.ProcessMSI(Arg)
      If Boo1 = False Then
          MsgBox "Failed to process MSI file. See the log file in " & MC.UnpackPath & " for more information.", 64, "MSI Unpacker"
     Else
          Boo1 = MC.UnPackMSI()   
          If Boo1 = True Then  
             MsgBox "Successfully unpacked to: " & vbCrLf & MC.UnpackPath & ".", 64, "MSI Unpacker" 
          Else
             MsgBox "Error unpacking MSI file. See the log file in " & MC.UnpackPath & " for more information.", 64, "MSI Unpacker" 
          End If
     End If
     
     Set MC = Nothing
     WScript.Quit
    
  '---------------------- begin MSI class -----------------------------     
Class MSIOps
 
Private FSO, SH, WI, DB, FolMSI, FolPack, FolData, DescPath, sBullet, sSpace, MSIPath, MSIType, sLine
Private TS, oFol, oFils, oFil, View, Rec                             '--various variables used repeatedly.
Private DicFiles, DicFolders, DicComps, DicFeat, DicReg, DicDesc      '-- dictionaries
Private HaveMSIData                                                   '--boolean value - whether the Class currently has a database Read in.
Private sCurrentFolderPath '-- get current folder .
'////////////// Public Functions /////////////////////////////////////


Public Property Let CurrentFolderPath(sPathCur1)  
    '-- HTA calls this at load. It gets window.location and cleans up that path as a way to get current folder path.
  sCurrentFolderPath = sPathCur1
End Property

Public Property Get UnpackPath()
  UnPackPath = FolMSI
End Property
'----------------- Read out the MSI database from file sPath and extract any internal CAB files.
'-- This Function calls Private LoadUpData, which confirms file existence and extension,
'-- Then processes database, filling Dic objects With info. about package.
'-- None of the several database functions will make this Call fail. It fails only If
'-- file does Not exist or If extension is Not MSI/MSM.
Public Function ProcessMSI(sPath)
  Dim Boo
     On Error Resume Next
    Boo = LoadUpData(sPath)
      If (Boo = False) Then
         ProcessMSI = False  '--  "Error processing file. Make sure Path is correct."
         Exit Function
      End If
  ProcessMSI = True
End Function

'---- Unpack MSI file. returns False If no msi file/database is currently Read into the Class. 
Public Function UnPackMSI()
  If (HaveMSIData = False) Then 
     UnPackMSI = False
     Exit Function
  End If
    DoFullUnPack '--Run Private Sub to unpack msi file.
  UnPackMSI = True  
End Function

'--Get a descriptive log of msi contents without unpacking:
Public Function WriteMSIDesc()
  If (HaveMSIData = False) Then 
      WriteMSIDesc = False
      Exit Function
  End If
    WriteLogFile 0
    WriteMSIDesc = True
End Function



'---<<<<<<<<< minor Public functions >>>>>>>>>>>>>> -------------------
'__________________________________________________
'--Get list of components in feature, returns array(1) With a(0)=description  a(1) = bullet-divided component list.
'--returns True If Afeat is loaded With array.
Public Function GetFeatureInfo(sFeat, AFeat)
   On Error Resume Next
     GetFeatureInfo = False
      If (HaveMSIData = False) Then Exit Function
    If DicFeat.exists(sFeat) Then
         AFeat = DicFeat.item(sFeat)
         GetFeatureInfo = True
    End If     
End Function

Public Function GetFeatureList(Ak)
   On Error Resume Next
     GetFeatureList = False
      If (HaveMSIData = False) Then Exit Function
    Ak = DicFeat.keys
     GetFeatureList = True
End Function

'__________________________________________________
'-- return array: acomp(0) = component folder. acomp(1) = bullet-divided file Path list. acomp(2) = bullet-divided file ID list. 
'-- acomp(3) = VBCrLf-divided reg.setting list With bullet-divided parts.
Public Function GetComponentInfo(sComp, AComp)
  Dim A2, s1, sFol, AFils, AKeys, i, sA1, sA2, sFils, sFilsID, ARet(3)
  On Error Resume Next
     GetComponentInfo = False
         If (DicComps.exists(sComp) = False) Then Exit Function
    A2 = DicComps.item(sComp)
    s1 = A2(0) '-- folder id For this comp.
        If (DicFolders.exists(s1) = False) Then Exit Function
    sFol = DicFolders.item(s1)
     AKeys = DicFiles.keys
       For i = 0 to UBound(AKeys)
          AFils = DicFiles.item(AKeys(i))
          sA2 = AFils(1)
           If (sA2 = sComp) Then '--If this file goes With the respective component....
               sFils = sFils & sFol & "\" & AFils(0) & sBullet  '--comp. folder Path + real name: full Path of file.
               sFilsID = sFilsID & AKeys(i) & sBullet           '--file ID
           End If
      Next
      
         '--  If (sFils = "") Then Exit Function      '--NOTE: 1-1-04. commented to prevent problem With Function returning False
                                                               '-- when component is only reg. settings.
      ARet(0) = sFol  '-- folder Path
      ARet(1) = sFils  '--list of file paths.
      ARet(2) = sFilsID '--list of file ids.
        If DicReg.exists(sComp) Then
           ARet(3) = DicReg.item(sComp)
        Else
           ARet(3) = ""
        End If
      AComp = ARet
      GetComponentInfo = True
End Function

'-- return array of component names:
Public Function GetComponentList(Ak)
   On Error Resume Next
     GetComponentList = False
      If (HaveMSIData = False) Then Exit Function
    Ak = DicComps.keys
     GetComponentList = True
End Function

'-- return component and feature associated With file. If file Not matched, return " ". If featrue Not matched, return " ".
'-- If no database loaded return "null"
Public Function FindFileComp(sFileName)
Dim AKeys, A2, i, s, sCL, ARet(1)
  On Error Resume Next
      If (HaveMSIData = False) Then 
           ARet(0) = "null"
           FindFileComp = ARet '--return "null" If Nothing is loaded.
           Exit Function
      End If    
     s = sSpace 
    AKeys = DicFiles.keys
       For i = 0 to UBound(AKeys)
          A2 = DicFiles.item(AKeys(i))
          If (UCase(A2(0)) = UCase(sFileName)) Then
              s = A2(1)
              Exit For
          End If    
       Next
  
   ARet(0) = s
        If (s = sSpace) Then      '--return " " If file Not found.
             FindFileComp = ARet
             Exit Function
        End If
  
     ARet(1) = sSpace
  AKeys = DicFeat.keys   '--find feature:
   For i = 0 to ubound(AKeys)
     A2 = DicFeat.item(AKeys(i))
     sCL = A2(1)
       If (InStr(1, sCL, (s & sBullet)) <> 0) Then  
           ARet(1) = AKeys(i)  ' comp + bullet + feature
           Exit For
       End If
   Next
     FindFileComp = ARet 
         
End Function

'--return summary desc. of package.
Public Function GetSummaryInfo()
   On Error Resume Next
      If (HaveMSIData = False) Then 
          GetSummaryInfo = ""
          Exit Function
     End If
       GetSummaryInfo = DicDesc.item("summary")     
End Function

'--return reg. list For component, separated by VBCrLf
Public Function GetCompRegList(sComp)
  Dim sReg
    On Error Resume Next
      If (HaveMSIData = False) Then 
          GetCompRegList = ""
          Exit Function
      End If
     
   If DicReg.exists(sComp) Then
      sReg = DicReg.item(sComp)
   Else
      sReg = ""
   End If
      GetCompRegList = sReg
End Function

'--return all reg. settings.
Public Function GetAllRegList()
Dim sReg, AKeys, i
   On Error Resume Next
      If (HaveMSIData = False) Then 
          GetAllRegList = ""
          Exit Function
      End If
     
AKeys = DicReg.keys
  For i = 0 to UBound(AKeys)
      sReg = sReg & DicReg.item(AKeys(i))
  Next
      GetAllRegList = sReg
End Function

'--process registry data into shell.RegWrite-ready lines.
'-- this Function takes a string composed of any number of utility-processed reg strings,
'-- VBCrLf-delimited With bullet-delimited data, and creates RegWrite-style strings from them.
'-- it returns VBCrLf-delimited strings like: 
'  SH.RegWrite "HKCR\CLSID\XXXXX-XXXX-XXXX-XXXXXXXX\InProcServer32\", "C:\Windows\System\SomeLib.DLL", "REG_SZ"
Public Function TranslateRegStrings(sRegIn)
Dim AReg, sPrepped, i, s1
   On Error Resume Next
    If sRegIn = "" Then
       TranslateRegStrings = ""
       Exit Function
    End If
  AReg = Split(sRegIn, vbCrLf)
    For i = 0 to UBound(AReg)
       s1 = Trim(AReg(i))
      If (s1 <> "") Then sPrepped = sPrepped & (PrepRegString(AReg(i)) & vbCrLf)
    Next  
  TranslateRegStrings = sPrepped
End Function

'---------- write a file:
Public Sub WriteData(sFil, sData)
 On Error Resume Next
    Set TS = FSO.CreateTextFile(sFil, True)
       TS.Write sData
       TS.Close
   Set TS = Nothing
End Sub

'-- add to a file:
Public Sub AppendData(sFil, sData)
 On Error Resume Next
    If FSO.FileExists(sFil) = False Then
         Set TS = FSO.CreateTextFile(sFil, True)
         Set TS = Nothing
    End If
      Set TS = FSO.OpenTextFile(sFil, 8)
          TS.Write sData
          TS.Close
      Set TS = Nothing
End Sub

Public Function GetFileList(SRet)
Dim AKeys, i, A2, s1
On Error Resume Next
      If (HaveMSIData = False) Then 
          GetFileList = False
          Exit Function
      End If
 AKeys = DicFiles.keys
      For i = 0 to UBound(AKeys)
        A2 = DicFiles(AKeys(i))
        s1 = s1 & A2(0) & vbCrLf   '--file name.
      Next
  SRet = s1  
   GetFileList = True
End Function

  '-- returns feature parent ID of given feature ID.
Public Function GetFeatureParent(sFeature)
  Dim A1
    On Error Resume Next
     A1 = DicFeat(sFeature)
     GetFeatureParent = A1(2)
End Function
  
'-- returns conditional install string For component.
Public Function GetCompCondition(sComp)
   Dim A1
    On Error Resume Next
     A1 = DicComps(sComp)
    GetCompCondition = A1(1)  
End Function

'//////////////////// End Public Functions ///////////////////////////

Private Sub Class_Initialize()
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set SH = CreateObject("WScript.Shell")
    sBullet = Chr(149)
    sSpace = " "
    HaveMSIData = False
 End Sub
          
Private Sub Class_Terminate()
     Set FSO = Nothing
     Set SH = Nothing
     Set DB = Nothing
     Set WI = Nothing
     Set DicFiles = Nothing  '--holds mapping of files. key-file ID  item(0)-file real name. item(1)-related component ID 
     Set DicFolders = Nothing '--holds mapping of folders needed. key-Directory ID.  item-local Path represented by that ID.
     Set DicComps = Nothing  '-- holds mapping of components. key-component ID. item-associated folder.
     Set DicFeat = Nothing    '-- holds mapping of features: key-feature name. item(0)-feature desc. item(1)-bullet-divided feature component list.
     Set DicReg = Nothing '--holds mapping of registry settings. key-component id  item- string of VBCrLf-divided registry settings With bullet-divided parts.
     Set DicDesc = Nothing '--holds log info. For other operations.
End Sub

Private Sub ResetAll()  '-- reset all objects when done With an MSI file:
     MSIPath = ""
     HaveMSIData = False
     Set DicFiles = Nothing
     Set DicFolders = Nothing
     Set DicComps = Nothing
     Set DicFeat = Nothing   
     Set DicReg = Nothing
     Set DicDesc = Nothing
     Set DB = Nothing
     Set WI = Nothing
     
    Set DicFiles = CreateObject("Scripting.Dictionary")
    Set DicFolders = CreateObject("Scripting.Dictionary")
    Set DicComps = CreateObject("Scripting.Dictionary")
    Set DicFeat = CreateObject("scripting.dictionary")
    Set DicReg = CreateObject("Scripting.Dictionary")
    Set DicDesc = CreateObject("Scripting.Dictionary")
    Set WI = CreateObject("WindowsInstaller.Installer")
End Sub



'--------------------------------------------------------------
' Function to create objects and create folder paths, based on file Path.
'-- no extensive error trapping here. If the file exists it's processed. If
'-- it doesn't exist the Function returns False.
Private Function LoadUpData(sFilePath)
Dim Pt1, sExt
  On Error Resume Next
       If (FSO.FileExists(sFilePath) = False) Then
           LoadUpData = False
           Exit Function
       End If
     sExt = UCase(Right(sFilePath, 3)) 
     
  If (sExt = "MSI") Then
     MSIType = 0                    '--MSI file.
  ElseIf (sExt = "MSM") Then
     MSIType = 1                    '-- merge module.
  Else
      LoadUpData = False
       Exit Function
  End If
  Pt1 = instrrev(sFilePath, "\")
  FolMSI = left(sFilePath, Pt1 - 1)   '--folder Path of MSI file.
  FolPack = FolMSI & "\Unpacked"      'folder to unpack MSI data to.
  FolData = FolPack & "\MSI_Utility_Data"           ' folder For holding MSI raw data during operation.
  DescPath = FolMSI & "\Program Description.txt" '-- Path of description file - needed later.    
  sLine = "______________________________"
  
  ResetAll
  Set DB = WI.OpenDatabase(sFilePath, 0)  '-- database object.

  '--Run functions to load msi database info:
     GetSummary           '--Get summary info and add to DicDesc.
    '    GetProductCode         '-- Get product code to use For working out hierarchy of products. 
     GetCABs                '--list CABs, extract If necessary, and add to DicDesc.
     ListFiles                 '-- Get list of files into DicFiles.
     SortFolders            '-- sort out the Directory table to map folder paths.
     SortComps             '-- sort out componnets into DicComps.
     DelineateFeatures   '-- Get feature descriptions With their respective component lists.
     CollectRegSets       '-- Get registry settings.
     MSIPath = sFilePath
     HaveMSIData = True
LoadUpData = True
End Function

'-- unpack MSI file currently Read into Class:

Private Sub DoFullUnPack()
    On Error Resume Next
    '-- make folders If necessary:
 
     If (FSO.FolderExists(FolPack) = False) Then
         Set oFol = FSO.CreateFolder(FolPack)
         Set oFol = Nothing
     End If
     If (FSO.FolderExists(FolData) = False) Then
         Set oFol = FSO.CreateFolder(FolData)
         Set oFol = Nothing
     End If
  
   ExtractAllCabs '--unpack CAB files to folder For processing.
       WScript.Sleep 500
    
   MakeMSIFolders  '-- make folder paths needed For package.
       WScript.Sleep 1000

   DistributeFiles  '--copy files into Path folders With real names.
   
   WriteLogFile 1    '--Write log file detailing the whole thing.
   
   FSO.DeleteFolder FolData, True '-- delete folder used to distribute files. ' 4/2014
End Sub


'888888888888888888888888888888888888888888888888888888888888888888888888888888
'///////////////////////// Workhorse functions to handle database /////////////////////////////////
'888888888888888888888888888888888888888888888888888888888888888888888888888888

Private Sub GetSummary() '--Get summary info from database.
  Dim SI, sSummary
   On Error Resume Next
   Set SI = DB.SummaryInformation(0) 
      sSummary = "Title: " & SI.Property(2) & vbCrLf
      sSummary = sSummary & "Subject: " & SI.Property(3) & vbCrLf
      sSummary = sSummary & "Author: " & SI.Property(4) & vbCrLf
      sSummary = sSummary & "Program Name: " & SI.Property(18) & vbCrLf
      sSummary = sSummary & "Creation Date: " & CStr(SI.Property(12)) & vbCrLf    
   Set SI = Nothing
        DicDesc.add "summary", sSummary
End Sub
' ___________________________________________________________

Private Sub GetCABs() '--check For CAB list in Media table. If there are internal CABs, extract them.
  Dim sCab, sCabs, sDesc, A1, i, sNewName
    On Error Resume Next
       '-- Get single CAB out If MSM file:
     If (MSIType = 1) Then
        GetTheCabOut "MergeModule.CABinet"
        Exit Sub
     End If
        
       Set View = DB.OpenView("SELECT `Cabinet` FROM Media")
       View.execute
            Do '--go through Media table and look For any records With a Cabinet value:
                Set Rec = View.Fetch
                If Rec is Nothing Then Exit Do
                sCab = Rec.Stringdata(1)
                   If (sCab <> "") And (InStr(1, sCabs, (sCab & ",")) = 0) Then
                       sCabs = sCabs & sCab & ","
                   End If   
            Loop 
             Set Rec = Nothing
       Set View = Nothing
    
             If (Len(sCabs) = 0) Then
                 DicDesc.add "cabs", "No CAB files listed as part of this install package."
                 DicDesc.add "cablist", ""
                 Exit Sub
             End If
    
           A1 = Split(sCabs, ",")
        For i = 0 to UBound(A1)
          sCab = A1(i)
            If (Left(sCab, 1) = "#") Then
               sCab = Right(sCab, (Len(sCab) - 1))  
               sNewName = GetTheCabOut(sCab)
               WScript.Sleep 500    '--make sure the cab file is there before calling extract.
               A1(i) = sNewName  '-- to drop # and also in Case cab was renamed when copied. (some cabs listed in media table are Not named CAB)
            End If   
        Next   
              sDesc = ""
              sCabs = ""
                '--check that all CABs are present:
          For i = 0 to ubound(A1)
              sCab = A1(i)
               If (sCab <> "") Then
                  If FSO.FileExists(FolMSI & "\" & sCab) = True Then
                     sDesc = sDesc & sCab & " - OK. File is present." & vbCrLf
                     sCabs = sCabs & FolMSI & "\" & sCab & ","
                  Else
                     sDesc = sDesc & sCab & " - CAB file needed but not found." & vbCrLf
                 End If 
               End If  
          Next     
          
        DicDesc.add "cabs", sDesc
        DicDesc.add "cablist", sCabs   '--save list of cab files For later extraction.
     
End Sub

' ___________________________________________________________
'-----------extract cab embedded in MSI file If necessary: ------------
'-- Surprisingly, this seems to work fine! It just reads out the CAB bytes and writes them to a "text" file.
Private Function GetTheCabOut(sCabName)
   Dim s, DLen, sFile
     On Error Resume Next
       Set View = DB.OpenView("SELECT `Name`,`Data` FROM _Streams WHERE `Name`= '" & sCabName & "'")
      View.execute
      Set Rec = View.Fetch
         If Rec is Nothing Then 
             Set View = Nothing
             GetTheCabOut = ""
             Exit Function
         End If    
           DLen = Rec.datasize(2)
           s = Rec.ReadStream(2, DLen, 2)
                   '--make sure it has a CAB extension For extraction purposes.
                   sFile = sCabName
               If (UCase(Right(sFile, 3)) <> "CAB") Then sFile = sFile & ".cab"
              Set TS = FSO.CreateTextFile(FolMSI & "\" & sFile, True, False)
                 TS.Write s
                 TS.Close
              Set TS = Nothing  
      Set Rec = Nothing
      Set View = Nothing
        GetTheCabOut = sFile
End Function


' ___________________________________________________________
'-----------extract files from all cabs when upacking MSI file:  

'-- ################################################################

'  -- ZIPEDIT-- CAB Extraction - THIS IS WHERE A ZIP PROGRAM  -WAS-  NEEDED IN THE OLD VERSION OF THE UNPACKER.
'-- YOU DO NOT NEED TO DO ANYTHING HERE IF jcabxmsi.exe IS IN THE SAME FOLDER.
'-- IF YOU *WANT* TO USE A ZIP PROGRAM TO UNPACK THE CAB(S) THEN FOLLOW DIRECTIONS BELOW.

'-- ___________________________________________________

'--- (THIS INFO. IS IRRELEVANT IF USING jcabxmsi.exe TO EXTRACT FROM CABS)

'  Below the command lines have been set up for PowerArchiver and 7-Zip zip programs.
'  How this works: You need to set up the command line for your zip program, including
' path. If you have PowerArchiver or 7-Zip the work is alreay done, but you'll still need
' to set the proper file path to the zip program executable.
'
'  Directions: Section A below is is path of zip program for command line. Make sure
' the path is valid on your system.
'
'  Section B and Seciton C are both the same. They contain full command lines for extraction.
'  The command line is basically:     "zip-program-exe-path" -e "cab-file-path" "folder-extract-to-path"
'  
'  Unfortunately, command lines are not standardized. The command lines below
'  are for PowerArchiver and 7-Zip, but if you use another zip program then you'll need to work out the
' specific command line for that program.

'-- ################################################################

Private Sub ExtractAllCabs()
Dim s, ACabs, i, Ret, sPath, Qt1
  On Error Resume Next
      Qt1 = Chr(34)
     sPath = Qt1 & sCurrentFolderPath & "jcabxmsi.exe" & Qt1 & " "
     
    
     '-- ---------- IGONORE THIS SECTION UNLESS EDITING TO USE A ZIP PROGRAM. -------------------
     
            '-- ==== Section A ======== SET PATH OF ZIP PROGRAM EXECUABLE HERE ==============
            '-- ======Only have one line here that is not commented.  ==================
              
                ' sPath = Qt1 & "C:\Program Files\PowerArchiver\POWERARC" & Qt1  '-- PowerArchiver.
                ' sPath = Qt1 & "C:\Program Files\7-Zip\7z.exe" & Qt1  '-- 7-zip.
      '--------------------------------------------------------------------------------------------
      
    If (MSIType = 1) Then    '-- MSM file.
        s = FolMSI & "\MergeModule.CABinet.cab"
         If FSO.FileExists(s) Then
                     '-- This command line is like so:   C:\MSI Unpacker\jcabxmsi.exe C:\unpack folder /V"
                     '-- The /V option is for "verbose". It will show error information if the extraction fails.
              Ret = SH.run(sPath & Qt1 & s &  "|" & FolData & " /V" & Qt1, , True)
            
                      '-- ----------ZIPEDIT-- IGONORE THIS SECTION UNLESS EDITING TO USE A ZIP PROGRAM. -------------------

                             '-- ==== Section B ======= EXTRACTION COMMAND. ONLY ONE LINE HERE CAN BE UNCOMMENTED. ==============
                        '   Ret = SH.Run(sPath & " -e " & Qt1 & s & Qt1 & " " & Qt1 & FolData & Qt1, , True)    '-- power Archiver.
                         ' Ret = SH.Run(sPath &  " e " & Qt1 & s & Qt1 & " -o" & Qt1 & FolData & Qt1, , True)   '-- 7-Zip.    
                     '--------------------------------------------------------------------------------------------
          End If
        Exit Sub
    End If
        '-- regular op For MSI files:
    s = DicDesc.item("cablist")
       If s = "" Then Exit Sub
  ' ACabs = Split(s, ",")
   ' For i = 0 to ubound(ACabs)
    '    s = ACabs(i)
        ' If Len(s) > 0 Then 
                      Ret = SH.run(sPath & Qt1 & FolMSI & "|" & FolData & Qt1, , True)
      '   End If
      
                     '-- ---------- ZIPEDIT--  IGONORE THIS SECTION UNLESS EDITING TO USE A ZIP PROGRAM. -------------------

                         '-- ==== Section C ======= EXTRACTION COMMAND. ONLY ONE LINE HERE CAN BE UNCOMMENTED. ==============
                        ' Ret = SH.Run(sPath  & " -e " & Qt1 & s & Qt1 & " " & Qt1 & FolData & Qt1, , True) '--  PowerArchiver.
                        ' Ret = SH.Run(sPath &  " e " & Qt1 & s & Qt1 & " -o" & Qt1 & FolData & Qt1, , True)   '-- 7-Zip.   
                    '--------------------------------------------------------------------------------------------
 '   Next   
End Sub

'-- #################### End CAB Extraction ops. ################################

'____________________________________________________________________
'-- Get list of all files.

Private Sub ListFiles()
Dim sName, sFil, a(1)
   Set View = DB.OpenView("SELECT `File`,`Component_`,`FileName` FROM `File`")
     View.execute
      Do
        Set Rec = View.Fetch
         If Rec is Nothing Then Exit Do
           sFil = Rec.stringdata(1) '--file ID
           sName = Rec.stringdata(3) '--filename
           sName = DropShort(sName)
           If (sFil <> "") And (DicFiles.exists(sFil) = False) Then
              a(0) = sName
              a(1) =  Rec.stringdata(2)  '-- component info.
              DicFiles.add sFil, a
                 '--once the cabs are unpacked these files can be looked For and
                 '--their names can be converted back to normal While assigning
                 '--them by component.
          End If
      Loop
      Set Rec = Nothing
    Set View = Nothing  
End Sub


' _________________________________________________________
'-- Read the Directory table and figure out the folder layout.
'--map all folder IDs to paths and record them in DicFolders as Key/Item = ID/Path.

Private Sub SortFolders()
Dim CFols, sID, sPar, sPar2, sDef, sPath, sPath2, sFolMainID, BooTarg
Dim Pt2, iCount, i2, i3, i20, iDone
On Error Resume Next
   BooTarg = False
   Set CFols = New FolderSort
    Set View = DB.OpenView("SELECT `Directory`,`Directory_Parent`,`DefaultDir` FROM `Directory`")
     View.execute
      Do
        Set Rec = View.Fetch
         If Rec is Nothing Then Exit Do
           sID = Rec.stringdata(1) '--folder ID
           sPar = Rec.stringdata(2) '-- parent dir.
           sDef = Rec.stringdata(3) '--def. directory.
           Pt2 = InStr(sDef, ":")
               If (Pt2 <> 0) Then sDef = Left(sDef, (Pt2 - 1)) '-- If DefDir is listed as target:source just get target.
           Pt2 = InStr(sDef, "|")
               If (Pt2 <> 0) Then sDef = Right(sDef, (Len(sDef) - Pt2))
            
           CFols.AddItem sID, sPar, sDef, ""                                          
      Loop    
        Set Rec = Nothing
        Set View = Nothing  

     
  '-- Next, find the top-level folder, if possible.
  
    iCount = CFols.ItemCount
    For i2 = 1 To iCount
         If BooTarg = True Then Exit For
       sID = CFols.GetIDByIndex(i2)
       sPar = CFols.GetParentByIndex(i2)
         If (sID = sPar) Or (Len(sPar) = 0) Then 'if parent is "" or matches ID then it may be targetdir.
                                                 'have to check whether any folders have this as parent.
           For i3 = 1 To iCount
             sPar2 = CFols.GetParentByIndex(i3)
             If (sPar2 = sID) And (i2 <> i3) Then       '-- look for at least 1 case where parent folder matches ID.
                 sFolMainID = sID   '-- This is to weed out dummy errors like two entries with blank
                 CFols.SetPathByIndex i2, FolPack
                 BooTarg = True                       '-- parent value or a table with a "TARGETDIR" ID where "TARGETDIR" is not actually the install directory.
                 Exit For
             End If
           Next
         End If
      Next
        '-- assuming there's a top-level folder, work out paths for all direct subfolders.
      If BooTarg = True Then
         For i3 = 1 To iCount
           sPar = CFols.GetParentByIndex(i3)
            If (sPar = sFolMainID) Then
                    sDef = CFols.GetDefDirByIndex(i3)
                 If sDef = "." Then
                      CFols.SetPathByIndex i3, FolPack
                 Else
                      CFols.SetPathByIndex i3, FolPack & "\" & sDef
                 End If    
            End If
         Next
      End If

       '-- resolve any "." defdir entries to parent directory. A value of "."
      ' as the target value in DefDir, in addition to there having been a src value,
      ' means this folder is actually the same
      ' as the parent folder. So check for "." and if found, set the DefDir
      ' value to that of the parent folder. But if defdir is simply "." it seems
      ' to indicate that ID is a directory variable, even though that's not in the help.
      ' example:  CommonFiles   TARGETDIR   .
      ' the only sensible way to translate that is to put a folder named CommonFiles in the unpack folder.
      ' in other words:   .:xyz - id is parent.   .  ID is defdir folder name.
      iDone = 0
  Do Until iDone = 3 '-- have to repeat this as some folders may be children of children.
    For i2 = 1 To iCount '-- resolve any "." def folder to parent folder.
        sDef = CFols.GetDefDirByIndex(i2)
       If sDef = "." Then  'for any entry where def folder is "."
         sPar = CFols.GetParentByIndex(i2)   ' get the parentfolder value
          For i3 = 1 To iCount
            sID = CFols.GetIDByIndex(i3)  '...then compare parent folder to sID.
            If sID = sPar Then       ' if directory of this entry matches directory_parent of entry with ".", then replace "." with this defdir.
               sDef = CFols.GetDefDirByIndex(i3)
               CFols.SetDefDirByIndex i2, sDef
               sPar2 = CFols.GetParentByIndex(i3)
               CFols.SetParentByIndex i2, sPar2
               Exit For
            End If
          Next
       End If
     Next
      iDone = iDone + 1
   Loop
   
      '-- Now all defdir entries should be valid names.
      '-- If there's no top-level folder ID found then set it to TARGETDIR---
      '-- and assign that with the path of the unpack folder.
      
       If BooTarg = False Then
         sFolMainID = "TARGETDIR---"
         CFols.AddItem sFolMainID, "", "", FolPack
       End If
       
    '-- at this point there should be a path for a top-level folder,
    '-- either targetdir or sfolmain (unpack folder or main folder in unpack folder)
    '-- Now go through items to fill in paths.
  Do
      iDone = 0
      i20 = i20 + 1
    For i2 = 1 To iCount  '-- for each folder...
        sPath = CFols.GetPathByIndex(i2) '-- see if there's a path yet.
      If Len(sPath) = 0 Then  '-- if not....
        iDone = 1 'set flag that all paths have not been found.
        sPar = CFols.GetParentByIndex(i2)  ' get the parent folder id and check it against directory ids.
              '  If sPar = sFolMainID Then.... not needed. All direct child folders under top level are already done.
            For i3 = 1 To CFols.ItemCount
              sID = CFols.GetIDByIndex(i3)
               If (sID = sPar) Then '-- if a match is found it's the parent folder. check to see if that has a path yet.
                  sPath2 = CFols.GetPathByIndex(i3)
                  sDef = CFols.GetDefDirByIndex(i2)
                  If Len(sPath2) > 0 Then  'if there's a path for parent folder then build this path.
                     sPath = sPath2 & "\" & sDef   'this just found a directory value to match the parent value,
                     CFols.SetPathByIndex i2, sPath 'then checked that for a path. If a path is there the path for
                     Exit For                   ' the current item is built and saved, until all paths are done.
                  End If
                End If
             Next
        End If
     Next
       If (iDone = 0) Or (i20 > 100) Then Exit Do
   Loop
   
    For i2 = 1 To iCount  '-- finally, add all ids/paths to DicFolders.
       sID = CFols.GetIDByIndex(i2)
       sPath = CFols.GetPathByIndex(i2)
       DicFolders.add  sID, sPath
    Next
    
         Set CFols = Nothing
  End Sub

'______________________________________________________
'-- Set up component dictionary. key = comp. id. item(0) = directory ID. item(1) = conditional install data.
Private Sub SortComps()
  Dim sComp, sDir, sCond, AComp(1)
    On Error Resume Next
   Set View = DB.OpenView("SELECT `Component`, `Directory_`, `Condition` FROM `Component`")
     View.execute
      Do
        Set Rec = View.Fetch
            If Rec is Nothing Then Exit Do
         sComp = Rec.stringdata(1) '--component ID.
         sDir = Rec.stringdata(2) '-- directory ID.
         sCond = Rec.stringdata(3) '-- conditional data.
           If (sComp <> "") And (DicComps.exists(sComp) = False) Then
                AComp(0) = sDir
                  If (Len(sCond) = 0) Then sCond = "none"
                AComp(1) = sCond
              DicComps.add sComp, AComp
           End If
      Loop
      Set Rec = Nothing
    Set View = Nothing  
End Sub

'______________________________________________________
'-- Set up features dictionary:   key-feature name.  item(0)-feature desc. 
'-- item(1)-bullet-divided list of components. item(2) - feature parent, If any.
Private Sub DelineateFeatures()
  Dim Ft, sInfo, sDes, AKeys, i, a(2), A2, sParentF
    On Error Resume Next
       If MSIType = 1 Then Exit Sub  '--no featrues For MSM.
       
    Set View = DB.OpenView("SELECT `Feature`, `Feature_Parent`, `Title`,`Description` FROM Feature")
     View.execute
      Do
        Set Rec = View.Fetch
            If Rec is Nothing Then Exit Do
         Ft = Rec.stringdata(1) '--feature name.
           If (Ft <> "") And (DicFeat.exists(Ft) = False) Then
               sInfo = "Title: " & Rec.stringdata(3) & vbCrLf '-- feature title.
               sInfo = sInfo & "Description: " & Rec.stringdata(4) & vbCrLf  & "Components -" & vbCrLf'-- feature description.
               a(0) = sInfo
               a(1) = ""
                 sParentF = Rec.stringdata(2) '-- feature parent.
                     If Len(sParentF) = 0 Then sParentF = "none"  
               a(2) = sParentF   
               DicFeat.add Ft, a
           End If
      Loop
      Set Rec = Nothing
    Set View = Nothing  
    
      '--Then Get the components that go With features. This goes through the
      '-- FeatureComponents table and adds components to list For Each feature
      '-- so that it can all be written to the program info. file:
      
    Set View = DB.OpenView("SELECT `Feature_`,`Component_` FROM FeatureComponents")
     View.execute
      Do
         Set Rec = View.Fetch
            If Rec is Nothing Then Exit Do
              Ft = Rec.stringdata(1) '--feature name.
           If DicFeat.exists(Ft) Then
              A2 = DicFeat.item(Ft)
              sInfo = A2(1)
              sInfo = sInfo & Rec.stringdata(2) & sBullet
              A2(1) = sInfo
              DicFeat.item(Ft) = A2
           End If
      Loop
      Set Rec = Nothing
    Set View = Nothing  
  
End Sub
' ______________________________________________________
'---Get the Registry settings.
Private Sub CollectRegSets()
Dim sReg, a(4), sRegRec, SRet
 On Error Resume Next
    Set View = DB.OpenView("SELECT `Registry`,`Root`,`Key`,`Name`,`Value`,`Component_` FROM Registry")
     View.execute
      Do
        Set Rec = View.Fetch
            If Rec is Nothing Then Exit Do
              sReg = Rec.stringdata(1) '--registry entry ID.
             If DicReg.exists(sReg) = False Then
                a(0) = Rec.integerdata(2) '--root
                a(1) = Rec.stringdata(3)  '-- key
                a(2) = Rec.stringdata(4)  '-- name
                a(3) = Rec.stringdata(5) '--value
                a(4) = Rec.stringdata(6)  '--component
                SRet = AddRegVal(a)
                 sRegRec = sRegRec & a(4) & "  --  " & SRet & vbCrLf
             End If
        Loop
      Set Rec = Nothing
    Set View = Nothing   
      
    DicDesc.add "registry", sRegRec
   
 End Sub

'_____________________________________
'-- turn reg table data into a usable string and store it by component.
Private Function AddRegVal(A2)
  Dim i, sReg, sName1, sVal, sType, sComp, sRegAll
    On Error Resume Next
   i = A2(0)
  Select Case i
    Case -1
       sReg = "UM\"
    Case 0
      sReg = "HKCR\"
    Case 1
      sReg = "HKCU\"
    Case 2
       sReg = "HKLM\"
    Case 3
       sReg = "HKU\"
  End Select
     sName1 = A2(2) 
     If (sName1 = "") Then sName1 = sSpace '-- put space If def. value (null)
     sReg = sReg & A2(1) & sBullet & sName1 & sBullet
     
   '-- Get value and check For type:
     sType = "SZ"
     sVal = A2(3)
   If (Len(sVal) > 1) Then
      If Left(sVal, 1) = "#" Then
        Select Case Left(sVal, 2)
           Case "#x"
               sType = "B"
           Case "#%"
               sType = "XS"
           Case  "##"
              sType = "SZ"
           Case Else
              sType = "D"      
        End Select
            '--strip off hash signs:
          sVal = Right(sVal, (Len(sVal) - 1))
          If (Left(sVal, 1) = "#") Then sVal = Right(sVal, (Len(sVal) - 1))      
      End If
   End If
   If (sVal = "") Then sVal = sSpace
     
 sReg = sReg &  sVal & sBullet & sType 
      '-- ( example:  "HKLM\Software\Microsoft\Windows\CurrentVersion•SystemRoot•C:\WINDOWS•SZ"  )
      '--                               Path                                                      value            data         type
     sComp = A2(4)
     sRegAll = ""
      '-- If there's already a key in DicReg For this comp. Then retrieve item. Otherwise, make a key.  
       If DicReg.exists(sComp) Then 
           sRegAll = DicReg.item(sComp)
       Else
           DicReg.add sComp, ""
       End If       
           '--save updated component list. Each dicitem can be split by VBCrLf to Get reg strings. Each reg string can be spit by bullet.
        sRegAll = sRegAll & sReg & vbCrLf
        DicReg.item(sComp) = sRegAll
  
   AddRegVal = sReg
End Function  

'--Do Function of ProcRegVals script. accept array of strings, return VBCrLf-delimited.
  '   (  HKLM\SOFTWARE\Microsoft\Key•Value•Data•Type  )
  '-- this Function is called For Each line from TranslateRegStrings. It builds the ProcRegVals script Function
  '-- into this Class so that the whole reg. string processing can be done from IE window.
Private Function PrepRegString(sRegStr)
Dim s1, s2, sDelim, sCom, A1, Q2, Val, Dat, sType, StrT
On Error Resume Next
  sDelim = "[~]"
  sCom = Chr(39) & "-- "
  Q2 = Chr(34)
     If (Left(sRegStr, 2) = "UM") Then 
          PrepRegString = sCom & sRegStr
          Exit Function
     End If  
        
     A1 = Split(sRegStr, sBullet)
         If (UBound(A1) <> 3) Then  
             PrepRegString = sCom & sRegStr
             Exit Function
         End If  
      
 s2 = "SH.RegWrite " & Q2 & A1(0) & "\"   '-- Set up beginning of string.
    Val = A1(1)
       If (Val = "+") Or (Val = "-") Or (Val = "*") Then  Val = " "      
    Dat = A1(2)  
    
         If (Val = " ") Then  '-- no value name. it's either a key or a default value.
              If (Dat = " ") Then  '--no value data. just Write the key. RegWrite won't Do that so Write "" to default value.
                  s2 = s2 & Q2 & ", " & Q2 & Q2      '--- "RegWrite "hklm\xxx\xxx\", ""     "
                  PrepRegString = s2
                  Exit Function
              Else
                  s2 = s2 & Q2 & ", " & Q2 & Dat & Q2            '--  "RegWrite "hklm\xxx\xxx\", "DefValueString"         "
                  PrepRegString = s2
                  Exit Function
             End If
         End If          
        
  '--at this point there's a value but may be no data.:
    If (Dat = " ") Then Dat = ""
   s2 = s2 & Val & Q2 & ", "   '--add value name to Path string.  'regwrite hklm\xxxx\xxxx\Val", '      "

  sType = A1(3)
    Select Case sType
      Case "B"
         StrT = Q2 & "REG_BINARY" & Q2
          PrepRegString = sCom & s2 & Dat & ", " & StrT
      Case "SZ"
           StrT = Q2 & "REG_SZ" & Q2
            If InStr(1, A1(2), sDelim) = 0 Then
               PrepRegString = s2 & Q2 & Dat & Q2 & ", " & StrT   
            Else
               PrepRegString = sCom & s2 & Q2 & Dat & Q2 & ", " & StrT
            End If   
      Case "XS"
        StrT = Q2 & "REG_EXPAND_SZ" & Q2
             If InStr(1, A1(2), sDelim) = 0 Then
                 PrepRegString = s2 & Q2 & Dat & Q2 & ", " & StrT
             Else
                 PrepRegString = sCom & s2 & Q2 & Dat & Q2 & ", " & StrT 
             End If   
      Case "D"
            StrT = Q2 & "REG_DWORD" & Q2
            PrepRegString = s2 & Dat & ", " & StrT
   End Select 
 
End Function

'-->>>>>>>>> Functions For unpacking MSI <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<---------------------------------

' ______________________________________________________
'-- cross reference the Dics to sort out files into folders.
'--files should all now be in the MSI_Utility_Data folder With IDs For names.
'-- so need to Get Each file, find it's component, find the folder
'-- For that component, and copy the file to folder With the right name.

Private Sub DistributeFiles()
  Dim sFil, sRealName, sFol, sFails, sComp, sFilPath, A2, A3, BooFail, stemp, oFols, oFol1, oFols1, sDel, ADel, i, sVers, i2
    On Error Resume Next
    Set oFol = FSO.GetFolder(FolData)
      Set oFils = oFol.Files
       If oFils.Count > 0 Then  '--if a CAB was embedded and unpacked.
             For Each oFil in oFils
                sFil = oFil.Name   
                 If DicFiles.exists(sFil) Then  '--If file is in file dictionary.....
                     A2 = DicFiles.item(sFil)
                     sRealName = A2(0)   '-- Get real file name.....
                     sComp = A2(1)        '-- ...and component that file belongs to.
                           BooFail = True
                             
                      If DicComps.exists(sComp) Then  '-- If component ID is in component dic...
                          A3 = DicComps.item(sComp)   '--Get folder For component.
                           sFol = A3(0)
                          '--this Next part distributes Each file to its folder, Then writes the string
                          '-- For Program Description log file. Example:         File: abc.DLL Version: 1.0.0.1
                          '                                                                   _C6584A3E3C2D4B1CB3A787C19D7F0EFF ->  \Systemfolder\ABC.DLL
                          
                            If DicFolders.exists(sFol) Then  '-- If Directory ID is in folders dic...
                                  sFilPath = DicFolders.item(sFol)  
                                  FSO.CopyFile FolData & "\" & sFil, sFilPath & "\" & sRealName, True
                                  FSO.DeleteFile FolData & "\" & sFil, True '4/2014 clean up distribution folder as unpacking proceeds.
                                  sVers = FSO.GetFileVersion(FolData & "\" & sFil)
                                  stemp = stemp & "File: " & sRealName & " Version: " & sVers & vbCrLf         
                                  stemp = stemp & sFil & " -> " & sFilPath & "\" & sRealName & vbCrLf & vbCrLf  '-- log files copied to program info. file.
                                  BooFail = False
                            End If
                      End If           
                   End If   
                           
                      '-- keep A list of any files Not successfully copied:
                     If (BooFail = True) Then sFails = sFails & sRealName & "-" & sFil & vbCrLf                     
               Next    
                  DicDesc.add "filecopy", stemp
                  
           Else  ' an MSI inside a self-executing CAB or other EXE that sorts the files itself.
             Dim AFils
              AFils = DicFiles.Items
              stemp = "The following files are part of the install but were not stored in an embedded CAB:" & vbCrLf & vbCrLf
              For i2 = 0 to UBound(AFils)
                A2 = AFils(i2)
                   sRealName = A2(0)   '-- Get real file name.....
                   sComp = A2(1)        '-- 
                   If DicComps.exists(sComp) Then  '-- If component ID is in component dic...
                       A3 = DicComps.item(sComp)   '--Get folder For component.
                       sFol = A3(0)
                        If DicFolders.exists(sFol) Then  '-- If Directory ID is in folders dic...
                             sFilPath = DicFolders.item(sFol)  
                             stemp = stemp & sFilPath & "\" & sRealName & vbCrLf
                        End If
                   End If
              Next   
                 DicDesc.add "filecopy", stemp  ' add to log in files copied section but these files are probably in the TEMP folder, 
                                                        ' unpacked And renamed by an EXE installer wrapper around the MSI.  
           End If           
                        
      Set oFils = Nothing
    Set oFol = Nothing    
      
        ' sFails = " Actual file name - File name in MSI_Utility_Data folder  -" & VBCrLf  & VBCrLf & sFails
       DicDesc.add "filefail", sFails   
       
        DeleteExtraFolders

End Sub

'--Delete unused folders. Called from DistributeFiles sub (above). This code updated 6-06. The former
'-- searched in the unpack folder for empty folders after unpacking was finished. The problem
'-- with that was that only the last folder in a hierarchy of empty folders was completely empty.
'-- This version uses the ist of folders that was created in the first place. It then loops
'-- through the list 4 times. Each time, for each item, it first checks whether the folder exists
'-- and then whether it's empty. If so, the folder is deleted. This method should be reasonably quick but will still
'-- delete empty folders down 4 levels.

Private Sub DeleteExtraFolders()
    Dim AFols3, iFols3, i4, UBFols, sFol3, oFol
  On Error Resume Next
         If (DicFolders.Count = 0) Then Exit Sub
    AFols3 = DicFolders.Items
    UBFols = UBound(AFols3)
      For i4 = 1 to 5
         For iFols3 = 0 to UBFols
           sFol3 = AFols3(iFols3)
            If (Len(sFol3) > 0) Then  '-- this part added in hopes of efficiency. setting all deleted indices to ""
                                             '-- allows for checking length of array(index) which should be much quicker than
                                             '-- checking whether a folder exists every time.
              If (FSO.FolderExists(sFol3) = True) Then
                  Set oFol = FSO.GetFolder(sFol3)
                     If (oFol.Files.count = 0) And (oFol.SubFolders.count = 0) Then
                         oFol.Delete True
                         AFols3(iFols3) = ""
                     End If
                  Set oFol = Nothing
              End If
           End If   
         Next
      Next         
End Sub

'______________________________________________________
'-- create all folders needed to unpack MSI file:
Private Sub MakeMSIFolders()
  Dim AFols, i, sFol
    On Error Resume Next
      If (DicDesc.exists("folders") = False) Then DicDesc.add "folders", ""
    AFols = DicFolders.items
       For i = 0 to UBound(AFols)
         sFol = AFols(i)
         MakeFolderPlus sFol       
      Next        
End Sub

'______________________________________________________
'-- Write Program Description file:
Private Sub WriteLogFile(iContent)   '-- 0 For list only. 1 If msi was unpacked and list includes files.
Dim sHead, sBod, AKeys, sKey, A1, i, s1, s2 
  On Error Resume Next
       If (HaveMSIData = False) Then Exit Sub
      
  If (FSO.FileExists(DescPath) = False) Then
       Set TS = FSO.CreateTextFile(DescPath, True)
         TS.WriteBlankLines 1
         TS.Close
       Set TS = Nothing
    End If
      
  sHead = "Descriptive information about " & MSIPath   '--summary desc.
  sBod = DicDesc.item("summary")
      WriteLogSection sHead, sBod
  
   If (MSIType = 0) Then   '--For MSI only.
              sHead = "Program CABs"                       '--cabs.
              sBod = DicDesc.item("cabs")
                  WriteLogSection sHead, sBod  
          
               sBod = ""
               AKeys = DicFeat.keys                          '-- features.
               For i = 0 to UBound(AKeys)
                  sKey = AKeys(i)
                     s1 = vbCrLf &  sLine & vbCrLf & "Feature Name: " & sKey & vbCrLf
                     A1 = DicFeat.item(sKey)
                     s1 = s1 & A1(0)
                     s2 = A1(1)
                     s2 = replace(s2, sBullet, vbCrLf)
                     s1 = s1 & s2
                    sBod = sBod & s1
               Next  
                 
                 sHead = "Feature Listing"
                 WriteLogSection sHead, sBod  
     End If                
    
    sHead = "Package Folder Paths"
    sBod = DicDesc.item("folders")
      '--snip FolPack string from folder paths:
      sBod = Replace(sBod, FolPack, "")
      sBod = Replace(sBod, sBullet, vbCrLf)
   
          WriteLogSection sHead, sBod  
    
   If (iContent = 1) Then  '-- unpacked.
      sHead = "Files copied: File ID -> Destination - "
      sBod = DicDesc.item("filecopy")
      sBod = Replace(sBod, FolPack, "")
          WriteLogSection sHead, sBod  
      sHead =  "Files without folder found: Actual file name - File name in MSI_Utility_Data folder  -"
      sBod = DicDesc.item("filefail")
          WriteLogSection sHead, sBod  
   End If
   
  sHead = "Registry settings: Component -- Setting - "    
  sBod = DicDesc.item("registry")
        WriteLogSection sHead, sBod  
End Sub

'-<<<<<<<<<<<< miscellaneous internal Function subs >>>>>>>>>>>>>>>>>>>------------------------------


'______________________________________________________
'-- Sub to make folder after checking to make sure it doesn't already exist.
Private Sub MakeFolder(sPath)
    On Error Resume Next
          If FSO.FolderExists(sPath)  Then Exit Sub
     Set oFol = FSO.CreateFolder(sPath) 
     Set oFol = Nothing
End Sub

'______________________________________________________
'-- Sub to make folder that may Not have an existing parent folder. also log list to DicDesc.
Private Sub MakeFolderPlus(sPath)
Dim Pt2, Pt3, s, sFol
    On Error Resume Next
          If FSO.FolderExists(sPath)  Then Exit Sub
        Pt2 = (instr(sPath, "\") + 1)
       Do
         Pt3 = InStr(Pt2, sPath, "\")
                If (Pt3 = 0) Then Exit Do
               s = Left(sPath, (Pt3 - 1))
             If FSO.FolderExists(s) = False Then
                 Set oFol = FSO.CreateFolder(s) 
                 Set oFol = Nothing
                 sFol = DicDesc.item("folders")
                 sFol = sFol & sBullet & s
                 DicDesc.item("folders") = sFol
             End If
          Pt2 = (Pt3 + 1)
       Loop
           If FSO.FolderExists(sPath) = False Then
               Set oFol = FSO.CreateFolder(sPath) 
               Set oFol = Nothing
                sFol = DicDesc.item("folders")
                sFol = sFol & sBullet & s
                DicDesc.item("folders") = sFol
           End If         
End Sub

'__________________________________________________________
'-- simple Function to drop short filenames where files or folders are listed in short|long format.
Private Function DropShort(sNom)
  Dim Pt2
   Pt2 = InStr(sNom, "|")  '--If filename is stored as short|long Then just take long.
   If (Pt2 <> 0) Then 
      DropShort = Right(sNom, (Len(sNom) - Pt2))
   Else
     DropShort = sNom
   End If     
End Function

'__________________________________________________________
'-- simple automation to Set up header when writing A section of data to log file:
Private Sub WriteLogSection(sHeader, sBody)
Dim TS1
 Set TS1 = FSO.OpenTextFile(DescPath, 8)
   TS1.Write sLine & vbCrLf & sHeader & vbCrLf & sLine & vbCrLf & vbCrLf & sBody & vbCrLf & vbCrLf
   TS1.Close
 Set TS1 = Nothing
End Sub

 End Class
 
'ooooooooooooooooooo End OF Class oooooooooooooooooooooooooooooooo

Class FolderSort
  Private ADir(), APar(), ADef(), APath()
 Private iCount, UB3

Private Sub BuildOut()
 iCount = iCount + 100
  ReDim Preserve ADir(iCount)
  ReDim Preserve APar(iCount)
  ReDim Preserve ADef(iCount)
  ReDim Preserve APath(iCount)
End Sub

  'store folder ID (key) and path (item)
Public Sub AddItem(sID, sPar, sDefName, sPath)
  On Error Resume Next
   If UB3 = iCount Then BuildOut 'if ubound reaches ubound of arrays then expand arrays.
  UB3 = UB3 + 1
  ADir(UB3) = sID
  APar(UB3) = sPar
  ADef(UB3) = sDefName
 APath(UB3) = sPath
End Sub

Public Function GetDefDirByIndex(LDex)
  On Error Resume Next
  GetDefDirByIndex = ADef(LDex - 1) 
End Function

Public Function GetIDByIndex(LDex)
  On Error Resume Next
  GetIDByIndex = ADir(LDex - 1)
End Function

Public Function GetParentByIndex(LDex)
  On Error Resume Next
  GetParentByIndex = APar(LDex - 1)
End Function

Public Function GetPathByIndex(LDex)
  On Error Resume Next
  GetPathByIndex = APath(LDex - 1)
End Function

Public Sub SetDefDirByIndex(iDex, sDef)
  On Error Resume Next
    ADef(iDex - 1) = sDef
End Sub

Public Sub SetParentByIndex(iDex, sPar)
  On Error Resume Next
    APar(iDex - 1) = sPar
End Sub

Public Sub SetPathByIndex(iDex, sPath)
  On Error Resume Next
    APath(iDex - 1) = sPath
End Sub

Public Function Exists(sID)
 Dim s 
 Dim i2
  On Error Resume Next
   Exists = False
   i2 = GetArrayIndexFromID(sID)
   If i2 > -1 Then Exists = True
End Function

Public Property Get Path(sID)
  Dim i2
  On Error Resume Next
    i2 = GetArrayIndexFromID(sID)
    If i2 > -1 Then Path = APath(i2)
End Property

Public Property Get ItemCount()
  On Error Resume Next
  ItemCount = UB3 + 1
End Property

Public Function GetArrayIndexFromID(sID)
  Dim i2
  On Error Resume Next
   GetArrayIndexFromID = -1
  If UB3 > -1 Then
    For i2 = 0 To UB3
      If ADir(i2) = sID Then
         GetArrayIndexFromID = i2
         Exit Function
      End If
    Next
  End If
End Function

Private Sub Class_Initialize()
   On Error Resume Next
  ReDim ADir(100) 
  ReDim APar(100)
  ReDim ADef(100) 
  ReDim APath(100)
  iCount = 100
  UB3 = -1
End Sub
End Class
