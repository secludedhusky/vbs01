'-- Display info script using IE ------------------------
'-- The IE Screen object provides access to display width in pixels,
'--display height in pixels, working area width, working area height
'-- and display Bits Per Pixel.
'-- In order to get at that info., the DisplayInfo class writes a webpage to
'--the Temp folder that contains an onload sub. An instance of IE is then
'--created and opens the webpage. The onload sub gets
'-- the display values and writes them to the page, alternating with
'--asterisks. (ex.: 800*600*800*560*24   )
'-- The script waits 1/2 second for the page to load and then gets
'-- the document.body.outertext, the page text.
'--Using the Split function, the values are put into an array and
'--then assigned to the GetInfo function parameters.
'--------------------------------------------------


Dim Display, i, w, h, aw, ah, bpp, s
   
       '--Create instance of DisplayInfo class:
  Set Display = New DisplayInfo
  
        '-- call GetInfo to return display data:
  i = Display.GetInfo(w, h, aw, ah, bpp)


          '--format data for message box:
     s = "Screen width: " & w & vbcrlf
     s = s & "Screen height: " & h & vbcrlf
     s = s & "Screen avail. width: " & aw & vbcrlf
     s = s & "Screen avail. height: " & ah & vbcrlf
     s = s & "Screen Bits Per Pixel: " & bpp

msgbox s, 64, "Display Info"

'---------------------------------------------------------------

'///////////////////////////////////////////////////////////////////
'/////// CLASS DisplayInfo ///////////////////////////////////////
'/////     GetInfo method returns:
'------- width and height of display in pixels.
'------- available width and height of display in pixels.
'------- display Bits Per Pixel.
'///////////////////////////////////////////////////////////////////////////////
'------Usage: Dim D
'------------ Set D = New DisplayInfo
'----------- i = D.GetInfo(width, height, availwidth, availheight, bitsperpixel)
'///////////////////////////////////////////////////////////////////////////////////////

Class DisplayInfo

Public Function GetInfo(vWidth, vHeight, vAvailW, vAvailH, vBitsPerPixel)
Dim FSO
Set FSO = CreateObject("Scripting.FileSystemObject")
  dim IE, s, TempFol, TS, a
On Error Resume Next
  TempFol = FSO.GetSpecialFolder(2)
  set TS = FSO.CreateTextFile(TempFol & "\sys.html")
  TS.WriteLine  "<HTML><HEAD><TITLE></TITLE><SCRIPT language=" & chr(34) & "vbscript" & chr(34) & ">"
TS.WriteLine "Sub doit()"
TS.WriteLine "Dim s"
TS.WriteLine "s = screen.width & " & chr(34) & "*" & chr(34)
TS.WriteLine "s = s & screen.height & " & chr(34) & "*" & chr(34)
TS.WriteLine "s = s & screen.availwidth & " & chr(34) & "*" & chr(34)
TS.WriteLine "s = s & screen.availheight & " & chr(34) & "*" & chr(34)
TS.WriteLine "s = s & screen.colordepth"
TS.WriteLine "document.write s"
TS.WriteLine "End Sub"
TS.WriteLine "</SCRIPT>"
TS.WriteLine "</HEAD>"
TS.WriteLine "<BODY onload=" & chr(34) & "doit()" & chr(34) & ">"
TS.WriteLine "</BODY></HTML>"
TS.Close
set TS = nothing
set FSO = nothing

Set IE = wscript.CreateObject("InternetExplorer.Application")
IE.visible = False
IE.Navigate "file:///" & TempFol & "\sys.html"
wscript.Sleep 500
s = IE.document.body.outertext
IE.quit
Set IE = nothing
if s <> "" then
      a = Split(s, "*")
      vWidth = a(0)
      vHeight = a(1)
      vAvailW = a(2)
      vAvailH = a(3)
      vBitsPerPixel = a(4)
      GetInfo = True
else
     GetInfo = False
end if
End Function

End Class