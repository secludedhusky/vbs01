'--COLUMNIZER SCRIPT.
'--------------------------------------------
'--This script will take a text file and reset the text in columns.

'--THERE'S ONE LIMITATION: The columns are made based on 
'--length of lines in terms of number of characters. The columns will only
'--line up when using a mono-spaced font such as Courier New. To line up
'--the text of other fonts would require knowing the character width of each character
'--in each font and then computing the actual length of lines.

'--The script starts by getting the file path. If file wasn't dropped on the script
'-- then a browsing window appears. The class that provides the browsing window
'--is at the end of the script.

'--Once there's a file the script asks for number of columns desired: 2, 3 or 4.
'--it then asks whether you want to columnize blocks of text.
'--Say, for example, you enter 3 for column number and 5 for text blocks.
'--the script will then make 3 columns. The first 5 lines of the file go in column 1.
'--lines 6-10 go in column 2, lines 11-15 go in column 3, lines 16-20 go
'--in column 1, and so on.
'--if you don't enter a number for block size the script will columnize by line.

'--The way it works: The file is read one line at a time. An array is created for
'--each column and each line is added to the array that represents its column.
'--while the line is being added, its length is checked and the longest length
'--in each column array is used to set column size.
'-- the script then opens a new file and writes one line at a time by reading
'--one line from each array, padding it to the respective column width,
'--and adding it to the next section.

'--There is a different sub used for block text because the arrays cannot be loaded in
'--order.

'--the new file will be created in the same path with the same name and "col" appended
'--to the beginning of the file name.

'--One use for this script: If you have a long list of items that you want to print out
'--you could make a multi-column list to use the printed page space more efficiently.
'-------------------------------------------------------------


Dim FSO, Obj, fPath, num, delim, r, ts, sFil, sFilB, NewPath, sMsg, BlockSize
Set FSO = CreateObject("Scripting.FileSystemObject")

 '-----------------get file path: ------------------------

If wscript.arguments.count <> 0 then
    fPath = wscript.arguments(0)
 else
    '--show file browsing window: ---------------
     Set obj = New FindFile
     fPath = Obj.Browse
end if

 if fso.FileExists(fPath) = false then
    wscript.Quit
 end if

'--set path for new file:

sFil = fso.GetFilename(fPath)
NewPath = left(fPath, (len(fPath) - len(sfil)))
sfilB = fso.GetBaseName(fPath)
NewPath = NewPath & "Col" & sFilB & ".txt"

'----------------get number of columns: --------------------------

num = 0
 do until num > 1 and num < 5
    num = inputbox("How many columns? 2, 3 or 4?", "Column Maker")
       if num > 1 and num < 5 then
            exit do
       end if
      r = msgbox("Entry must be 2, 3 or 4. Do you want to continue?", 33, "Column Maker")
        if r = 2 then
           wscript.Quit
       end if
 loop

'-----------------get number of lines intended to be in each column: ----------------------------------------

on error resume next
sMsg = "Normally this script will go line by line." & vbcrlf
sMsg = sMsg & "For example, with 2-columns line 1 will go in column 1, " & vbcrlf
sMsg = sMsg & "line 2 in column 2, line 3 in column 1, and so on." & vbcrlf
sMsg = sMsg & "If you want to arrange blocks of text enter the number" & vbcrlf
sMsg = sMsg & "of lines that should be treated as one column unit."

delim = inputbox(sMsg, "Column Maker")

     if delim = "" or delim < 2 or isnumeric(delim) = false then     'no delimiter string.
        Blocksize = 1
     else
        Blocksize = delim
     end if
     
     if Blocksize <> 1 then
       WriteColVersionMulti fPath, NewPath, num, BlockSize
     else
       WriteColVersion fPath, NewPath, num
     end if

msgbox "Done."

'--END OF SCRIPT ------------------------------------------------------


'*******************************************************************
'-------------------BEGIN SUB TO PROCESS FILE WITH MULTIPLE LINES IN EACH COLUMN ITEM. ----------------------------
'-------------------This sub will put BlockSize number of lines in 1st column, then
'-------------------BlockSize number of lines in 2nd column, etc., returning to column 1 after other columns 
'-------------------have been done.
'-------------------Each file line has to be built from lines in original file, combined with spaces.
'-------------------For multi-line blocks you need to read BlockSize # of lines into each array at a time
'------------------so that when the text is put back together the proper lines will line up with each other.

'*******************************************************************

sub WriteColVersionMulti(sPath, NPath, NumCols, BS) 
  dim i, i2, iLines, ANum1, ANum2, ANum3, ANum4, L1, L2, L3, L4, s, A1(), A2(), A3(), A4()
   
    set ts = fso.OpenTextFile(sPath, 1, False)
      
      '--read file one line at a time. Assign each line to it's column array and track the length of longest line:
  ANum1 = -1
  ANum2 = -1
  ANum3 = -1
  ANum4 = -1
     Do While Not ts.AtEndOfStream
       
          For i = 1 To Numcols  '-- i is number of columns. Do a read for each column and assign it to its array.
             If Ts.AtEndOfStream = True Then
                Exit For
             End If
                For iLines = 1 to BS                              '--within the For loop for loading one array item, read the requisite number
                       If Ts.AtEndOfStream = True Then    '--of lines.
                              Exit For
                       End If
                    s = ts.ReadLine
                    s = trim(s)
                      select case i
                             Case 1
                                 ANum1 = ANum1 + 1  '--ANum is used to increment array item numbers.
                                  Redim Preserve A1(ANum1)
                                  A1(ANum1) = s                
    
                                    If Len(s) > L1 Then
                                        L1 = Len(s)
                                     End If
       
                             Case 2
                                 ANum2 = ANum2 + 1 
                                 Redim Preserve A2(ANum2)
                                    A2(ANum2) = s       
                                  
                                   If Len(s) > L2 Then
                                      L2 = Len(s)
                                   End If
                        
                            Case 3
                                ANum3 = ANum3 + 1 
                                Redim Preserve A3(ANum3)
                                   A3(ANum3) = s       
                                   
                                   If Len(s) > L3 Then
                                      L3 = Len(s)
                                   End If
        
                            Case 4
                                  ANum4 = ANum4 + 1 
                                Redim Preserve A4(ANum4)                                  
                                       A4(ANum4) = s 
                                 
                                    If Len(s) > L4 Then
                                      L4 = Len(s)
                                   End If
                        End Select
                  Next
              Next
        Loop
   
   ts.Close
   Set Ts = Nothing

'--these numbers will now represent the char. length of the longest line
'--in each column. Add 5 to each for spacing:

L1 = L1 + 5
L2 = L2 + 5
L3 = L3 + 5
L4 = L4 + 5

  Set ts = FSO.CreateTextFile(NPath, True)

 On Error Resume Next

 '--build a file text line with one line from each array, padded to be 5 char. longer than longest entry in column.
 
    For I2 = 0 To Ubound(A1)
         S = A1(i2) & Space(l1 - Len(A1(i2)))
           
          If Ubound(A2) >= I2 Then 
             S = S & A2(i2) & Space(l2 - Len(A2(i2)))
          End If
        
          If Ubound(A3) >= I2 Then 
             S = S & A3(i2) & Space(l3 - Len(A3(i2)))
          End If
        
          If Ubound(A4) >= I2 Then 
             S = S & A4(i2) & Space(l4 - Len(A4(i2)))
          End If
            
         Ts.writeline S  '-- Write Formatted String To File.
     Next

   Ts.close
   Set Ts = Nothing
     
End Sub


'*******************************************************************
'-------------------begin Sub To Process File with single lines put into columns.
'-------------------This will assign each line to a column, repeating until the end.
 '---(line 1 goes to col. 1, line 2 goes to col. 2, etc. It then repeats so that if there
'--  are 2 columns line 3 will go to col. 1, line 4 to col. 2, line 5 to col. 1, etc  )
'********************************************************************
sub WriteColVersion(sPath, NPath, NumCols) 
 dim i, i2, ANum, L1, L2, L3, L4, s, A1(), A2(), A3(), A4()
   
    set ts = fso.OpenTextFile(sPath, 1, False)
      
      '--read file one line at a time. Assign each line to it's column array and track the length of longest line:
  ANum = -1
     Do While Not ts.AtEndOfStream
       ANum = ANum + 1
          For I = 1 To Numcols
             If Ts.AtEndOfStream = True Then
                Exit For
             End If
             s = ts.ReadLine
             s = trim(s)
               select case i
                      case 1
                           Redim Preserve A1(ANum)
                           A1(ANum) = s                
                              If Len(s) > L1 Then
                                 L1 = Len(s)
                              End If
                      case 2
                          Redim Preserve A2(ANum)
                          A2(ANum) = s                       
                            If Len(s) > L2 Then
                               L2 = Len(s)
                            End If
                      Case 3
                         Redim Preserve A3(ANum)
                         A3(ANum) = s
                            If Len(s) > L3 Then
                               L3 = Len(s)
                            End If
                      Case 4
                         Redim Preserve A4(ANum)                      
                         A4(ANum) = s
                             If Len(s) > L4 Then
                               L4 = Len(s)
                            End If
                 End Select
          Next
     Loop
   
   ts.Close
   Set Ts = Nothing

L1 = L1 + 5
L2 = L2 + 5
L3 = L3 + 5
L4 = L4 + 5

  Set ts = FSO.CreateTextFile(NPath, True)

 On Error Resume Next

 '--build a line with one line from each array, padded to be 5 char. longer than longest entry in column.
 
    For I2 = 0 To Ubound(A1)
         S = A1(i2) & Space(l1 - Len(A1(i2)))
           
          If Ubound(A2) >= I2 Then 
             S = S & A2(i2) & Space(l2 - Len(A2(i2)))
          End If
        
          If Ubound(A3) >= I2 Then 
             S = S & A3(i2) & Space(l3 - Len(A3(i2)))
          End If
        
          If Ubound(A4) >= I2 Then 
             S = S & A4(i2) & Space(l4 - Len(A4(i2)))
          End If
            
         Ts.writeline S  '-- Write Formatted String To File.
     Next

   Ts.close
   Set Ts = Nothing
     
end sub

'********************************************************************
'********************************************************************
'-----------------BEGIN CLASS BLOCK FOR FILE BROWSING WINDOW  -----------------------------------
'-- Use:         Set obj = New FindFile
'--                 var = obj.Browse
'-------------- var returns path of file selected. -----------------------------
'---------------------------------------------------------------------
Class FindFile
            Private fso, sPath1
          
      '--FileSystemObject needed to check file path:
           Private Sub Class_Initialize()
                Set fso = CreateObject("Scripting.FileSystemObject")
           end sub
          
           Private Sub Class_Terminate()
              Set FSO = Nothing
           End sub
  
      '-- the one public function in class:
       
           Public Function Browse()
              on error resume next
               sPath1 = GetPath
               Browse = sPath1
           end function

 Private Function GetPath()
     Dim Ftemp, ts, IE, sPath, sStatus

    '--Get the TEMP folder path and create a text file in it:

            Ftemp = fso.GetSpecialFolder(2)
            Ftemp = Ftemp & "\FileBrowser.html"
            set ts = fso.CreateTextFile(Ftemp, true)

 '--write the webpage needed for file browsing window:

            ts.WriteLine "<HTML><HEAD><TITLE></TITLE></HEAD>"
            ts.WriteLine "<BODY BGCOLOR=" & chr(34) & "#C7C7E2" & chr(34) & " TEXT=" & chr(34) & "#3B3B80" & chr(34) & ">"
                 ts.WriteLine "<script language=" & chr(34) & "VBScript" & chr(34) & ">"
                         '--sub for CANCEL button to catch click in script:
                     ts.WriteLine "sub butc_onclick()"
                        ts.WriteLine "status = " & chr(34) & "cancel" & chr(34)
                     ts.WriteLine "end sub"
                 ts.WriteLine "</script>"
            ts.WriteLine "<DIV ALIGN=" & chr(34) & "left" & chr(34) & ">"
            ts.WriteLine "<FONT FACE=" & Chr(34) & "arial" & Chr(34) & " SIZE=2>"
            ts.WriteLine "This script will convert text into multiple columns.<BR>"
            ts.WriteLine "First - Select a file.<BR>"
            ts.WriteLine "Next - The script will ask how many columns.<BR>"
            ts.WriteLine "Enter 2, 3 or 4 as the number of columns.<BR>"
             ts.WriteLine "Third - The script will ask whether you want to columnize text blocks.<BR>"
              ts.WriteLine "If nothing is entered in that input box the script will make a column <BR>"
               ts.WriteLine "row from each 2, 3 or 4 lines in the file. If a number is entered for <BR>"
               ts.WriteLine "block text the script will treat each text block as it would <BR>"
             ts.WriteLine "have treated a line. For example: If 5 is entered then 5 lines <BR>"
              ts.WriteLine "will be put into col. 1, the next 5 into col. 2, etc.<BR><BR>"
             ts.WriteLine "NOTE: Columns will only show properly in a monospaced font.<BR><BR>"
              ts.WriteLine "</DIV><DIV ALIGN=" & chr(34) & "center" & chr(34) & ">"
            ts.WriteLine "<FORM>"
                 
                     '--this is the file browsing box in webpage:
                 ts.WriteLine "<INPUT TYPE=" & chr(34) & "file" & chr(34) & "></input>"
                      '--this puts a bit of space between buttons:
              ts.WriteLine chr(38) & chr(35) & "160" & chr(59) & " " & chr(38) & chr(35) & "160" & chr(59)
                     '--this is the CANCEL button:
                 ts.WriteLine "<input type=" & chr(34) & "button" & chr(34) & " id=" & chr(34) & "butc" & chr(34) & " value=" & chr(34) & "CANCEL" & chr(34) & "></input>"
           ts.WriteLine "<BR>"
           ts.WriteLine "</FORM>"
           ts.WriteLine "</FONT></DIV>"
           ts.WriteLine "</BODY></HTML>"
           ts.Close
           set ts = nothing
               on error resume next

'--webpage is written. now have IE open it:

               Set IE = Wscript.CreateObject("InternetExplorer.Application")
                     IE.Navigate "file:///" & Ftemp
                     IE.AddressBar = false
                     IE.menubar = false
                     IE.ToolBar = false
                     IE.StatusBar = false
                     IE.width = 460
                     IE.height = 340
                     IE.resizable = false
                     IE.visible = true
 
 '--do a loop every 1/2 second until either:
'-- the browsing window value is a valid file path or 
'-- CANCEL is clicked (setting IE.StatusText to "cancel") or
'--IE is closed from the control box.

                     Do while IE.visible = true
                          spath = ie.document.forms(0).elements(0).value         '--get browsing text box value.
                          sStatus = IE.StatusText                                         '--get status text.

                             if fso.FileExists(spath) = true  or sStatus = "cancel" or IE.visible = false then
                                 exit do 
                             end if
                          wscript.sleep 500
                     Loop
                    
                IE.visible = false
                IE.Quit
                set IE = nothing
                  if fso.FileExists(spath) = true then
                      GetPath =  spath
                 else
                      GetPath = ""
                 end if   
  End Function

End Class