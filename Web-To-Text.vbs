'-- Web-To-Text will convert a webpage to text. Drop an HTML file on the script or double-click
 '-- and enter the full path.
'--  If the file was named, For example, index.html a file will appear in the
'-- same folder called index.txt. The URL text of links can be included or Not.

'--**the script starts by asking whether you want to save URL text. That means that any links
'--    in the page will also include the URL, written as :  url - link text. ex.: http:\\www.yahoo.com - Go To Yahoo

'--**Next the script gets the file and reads it into a textstream, dropping blank lines.

'--** Then the string is sent to StripIt Function to replace relevant tags, such as <BR>,
'-- and clean out all non-relevant tags. 
'--stripit goes through from one tag to the Next, getting the text between them.
'--    If it finds a link tag (HREF=) it can also Get the URL from that and insert it ahead of the link text seen on the webpage.

'-- **The string without tags is sent back and all ASCII codes are translated. 

'--** finally, the string is written to a new file with .TXT extension.

'-- The results from this script will vary depending on how the HTML page was
'--written. Ideally all vbCrLf should be removed before processing but that would
'--also require that all header tags, <PRE> tags, etc. be processed. This script
'--uses a simpler method that work fairly well in most cases.
 
'---------------------------------------------------------------------


Dim FSO, fHTMLPath, fNewPath, fdot, s1, s2, BoolURL, r, f1, Pt1, Pt2, PtSemiC, sTemp, sChar, sClean
Set FSO = CreateObject("Scripting.FileSystemObject")
On Error Resume Next

 
'--Get path of HTML file:

    If WScript.arguments.count = 0 Then
           fHTMLPath = InputBox("Enter the path of the file to process.", "Filepath")
     Else
           fHTMLPath = WScript.arguments.item(0)
     End If

If fHTMLPath = "" Then
  Set FSO = Nothing
  WScript.quit
End If

  If FSO.fileexists(fHTMLPath) = False Then
      MsgBox "The path is wrong. There's no such file.", 16, "Woops!"
      Set FSO = Nothing
      WScript.quit
  End If
  
'--create path For new file:

  fdot = instrrev(fHTMLPath, ".")
  fNewPath = (left(fHTMLPath, fdot)) & "txt"
 
  '----------------------------------find out whether to save url text For links.
 r = MsgBox("If this file contains links that you want to save, the link URL text can be included in the new file. Do you want to do that?", 36, "Save Link URLs?")
   If r = 6 Then
       BoolURL = True
  Else
       BoolURL = False
  End If
  
'--start by cleaning white space from file:

 Set f1 = FSO.OpenTextFile(fHTMLPath, 1)
  Do While f1.AtEndOfStream = False
       s1 = f1.ReadLine
        If s1 <> "" Then  '-------------take out space left by stripping HTML
           s2 = s2 & s1 & vbcrlf
        End If
          If f1.AtEndOfStream = True Then
             Exit Do
          End If
   Loop
  f1.Close
 Set f1 = Nothing   '---------s2 is now the file with empty lines removed. now convert spec. characters.
 
 '--Send cleaned string to StripIt to process tags:
   sClean = StripIt(s2)
   
 '-- now text has been cleaned of space and relevant tags have been translated.
 '-- replace common browser-specific codes If they were used:
 
  sClean = Replace(sClean, chr(38) & "nbsp;", " ")
  sClean = Replace(sClean, chr(38) & "quot;", chr(34))
  sClean = Replace(sClean, chr(38) & "amp;", chr(38))
  sClean = Replace(sClean, chr(38) & "gt;", ">")
  sClean = Replace(sClean, chr(38) & "lt;", "<")
  
   '--replace all ascii code strings written as "&#" + ascii code + ";"

 Pt2 = 1
  Do
    Pt1 = InStr(Pt2, sClean, "&#", 1)
        If Pt1 = 0 Then Exit Do
    PtSemiC = InStr(Pt1, sClean, ";", 1)
      If (PtSemiC > 0)  and (PtSemiC < (Pt1 + 6)) Then
         sTemp = Mid(sClean, (Pt1 + 2), (PtSemiC - (Pt1 + 2)))
         sChar = Chr(CInt(sTemp))
         sTemp = "&#" & sTemp & ";"
         sClean = Replace(sClean, sTemp, sChar)
      End If
     Pt2 = Pt1 + 1   
  Loop       
  
  '--last - to make up For Not processing some of the sacing tags,
  '-- check For excessive spacing:
  sTemp = vbcrlf & vbcrlf & vbcrlf & vbcrlf
  sChar = vbcrlf & vbcrlf
  sClean = Replace(sClean, sTemp, sChar)
  sTemp = vbcrlf & vbcrlf & vbcrlf 
  sClean = Replace(sClean, sTemp, sChar)
  
  '----write the new file in same folder as html file.
  
 Set f1 = FSO.createTextFile(fNewPath, True)
 f1.Write sClean
 f1.Close
Set f1 = Nothing
Set FSO = Nothing
 MsgBox "All done.", 0, "done"
 

'-----------Function StripIt - called after white space is cleaned out.---------------
'-------Replaces <BR>, <HR> and <P> tags.--------
'--cleans out irrelevant tags and gets links If BoolURL is True.

Function StripIt(sHTML)    '------Sub to strip HTML tags.
 Dim RightPt, LeftPt, s, LenT, StartPt, sTag, TagPt, QuotePt, sRep, sStripped
     
       '--replace all line break tags with carriage returns:
  sHTML = Replace(sHTML, "<BR>", vbcrlf)
  sHTML = Replace(sHTML, "</P>", vbcrlf)
  
       '--replace all horizontal line tags with lines:
   sRep = vbcrlf & "__________________________" & vbcrlf
    sHTML = Replace(sHTML, "<HR>", sRep)
    
      '--replace all paragraph starts:
   sRep = vbcrlf & vbcrlf & "      "   
      sHTML = Replace(sHTML, "<P>", sRep) 
      
  '--  Get all text between > and <.
  '-- If BoolURL = True Then Get text of links as well:
  
     StartPt = 1 '--point to start search.
     RightPt = 1 '--point where > is found.
     LenT = len(sHTML)
     
  Do While RightPt <> 0
      RightPt = instr(StartPt, sHTML, ">", 1)   '-------------find > tag. StartPt is the last < tag.
         If (RightPt = 0) or (RightPt >= LenT) Then Exit Do
       
      LeftPt = instr(RightPt, sHTML, "<", 1)  '----------this finds the Next < tag.
         If (LeftPt = 0) or (LeftPt >= LenT) Then Exit Do
      
   '------------------------------Get url text If that option chosen.
           If BoolURL = True Then
              sTag = mid(sHTML, StartPt, (RightPt - StartPt)) '--text within last tag.
              If (sTag <> "") Then  
                   TagPt = instr(1, sTag, "HREF=", 1) 
                       If (TagPt <> 0) Then  '-------------If it's a link tag Then....
                          TagPt = TagPt + 6  '--move search to beyond fist " after HREF=
                           QuotePt = instr(TagPt, sTag, chr(34), 1)
                           sTag = mid(sTag, TagPt, (QuotePt - TagPt)) '---------Get the link url.....
                           sStripped = sStripped & sTag & " - "    '------------and write it in the new file.
                       End If
                End If
            End If
  '------------------------------------------------
          If (LeftPt - (RightPt + 1)) > 0 Then   '------------If there's any text between tags write it to file.
             s = mid(sHTML, (RightPt + 1), (LeftPt - (RightPt + 1)))
              sStripped = sStripped & s
          End If
          
      StartPt = LeftPt
       If StartPt > (LenT - 2) Then
             Exit Do
       End If
   Loop
 
  StripIt = sStripped
 End Function