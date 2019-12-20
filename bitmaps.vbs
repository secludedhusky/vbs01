'--this script will make .bmp files display as pictures of what's inside instead of
'--as icons. It also provides for you to reverse the process by rewriting the script.

Dim sh, r, current, fso, fil
Set fso = CreateObject("Scripting.FileSystemObject")
Set sh = WScript.CreateObject("WScript.Shell")

r = msgbox("This will make all the .bmp files on your computer display in miniature version, instead of as file icons, when you view them in My Computer or Windows Explorer. Do you want to do that?", 36, "Showing Bitmaps")

If r = 7 Then  '--No was clicked.
   wscript.quit
else  
    '--get the current BMP setting and write it to a file:
  
    current = sh.RegRead("HKEY_CLASSES_ROOT\Paint.Picture\DefaultIcon\")
   Set fil = fso.CreateTextFile ("C:\bmpsettings.txt", true)
        fil.writeline "The old setting for .bmp files is* " & current & " *,without the asterisks. If you want to reverse the process then right click the bitmap.vbs file, click Edit, and replace %1 in line 22 with the old setting. Then save the file and run it." 
        fil.close
   set fil = nothing
   msgbox "The old setting has been saved in the file bmpsettings.txt on C drive, along with directions in case you want to reverse this process. You'll need bitmap.vbs, as well, to reverse it.", 64, "Preparing to change setting"
end if

gotoit = msgbox("Change bitmap file view setting now?", 36, "Change bitmap settings?")
   If gotoit = 7 then
        wscript.quit
   else
      '--write the registry setting to make bitmaps show as icons.
      '--The data is being written blank first to make sure it overwrites the old value.
            sh.RegWrite "HKEY_CLASSES_ROOT\Paint.Picture\DefaultIcon\", "", "REG_SZ"
            sh.RegWrite "HKEY_CLASSES_ROOT\Paint.Picture\DefaultIcon\", "%1", "REG_SZ"
       MsgBox "The setting has been changed.", 0, "All Set"
  End if



