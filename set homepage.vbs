'--this script will ask for a file path or URL and then set that
'--as your IE homepage. See below for directions to program the
'--script to set a given homepage automatically.

Dim sh, r, s, s1, sSp, fso
Set sh = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")


s1 = "Enter file path or URL for IE Home Page." & VbCrLf
s1 = s1 & "Samples of valid format:" & VbCrLf
s1 = s1 & "C:\homepage.html" & VbCrLf
s1 = s1 &  "www.someplace.com" & VbCrLf
s1 = s1 & "www.someplace.com\somepage.html"

sSp = chr(37) & "20"     '--used to format file path string because browser won't read spaces in path.

on error resume next

'---------------**************************************
'---------------------------------------------------------

'--NOTE: TO MAKE THIS SCRIPT WORK AUTOMATICALLY,
'--IN THE NEXT LINE SUBSTITUTE  inputbox(s1, "Enter home page path or URL")
'--WITH DESIRED HOME PAGE.
'--EX.: s = "www.google.com" or s = "C:\webpages\homepage.html"
'--YOUR HOME PAGE WILL THEN BE RESET SILENTLY WHEN YOU RUN THE SCRIPT.
'--TO PREVENT THE CONFIRMATION MESSAGE PUT AN APOSTROPHE IN FRONT
'--OF LAST LINE.

'------------------------------------------------------------
'-************************************************

 s = inputbox(s1, "Enter home page path or URL")

    if s = "" then
       wscript.Quit
    end if

'--check whether file or URL by looking for drive designation: -------------
       '--if file, substitute "%20" for spaces and add "file:///" at beginning: ---------------

   if mid(s, 2, 2) = ":\" then  '--it's a local file. -----------
           if fso.FileExists(s) = false then
              msgbox "File does not exist.", 64, "No such file"
              wscript.Quit
           end if
   
      s = Replace(s, chr(32), sSp, 1, -1, 1)
      s = "file:///" & s
  else
           '--it's a URL. add "http://" if necessary. -----------------

        if left(s, 7) <> "http://" then
           s = "http://" & s
        end if
  end if
Set fso = nothing
  '--write registry setting. This seems to work more dependably if you write a blank first. ---------------

sh.RegWrite "HKCU\Software\Microsoft\Internet Explorer\Main\Start Page", "", "REG_SZ"
sh.RegWrite "HKCU\Software\Microsoft\Internet Explorer\Main\Start Page", s, "REG_SZ"

Set sh = nothing
msgbox "Done. Home page is now: " & s, 64, "Start Page set"


