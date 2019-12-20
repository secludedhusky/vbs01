'--this script will change a registry setting to let you change the title on
'--Internet Exp. title bar. It's been updated to work in IE 5.1
Dim sh, v, r
Set sh = WScript.CreateObject("WScript.Shell")

v = Msgbox("This allows you to change the text on the title bar of Internet Explorer. Normally, at the top it prints the page that you're at and then either 'Microsoft Internet Explore'r or something your ISP has put there.{ You could open Internet Explorer now if you want to see.}", 1, "Little Tweak!")
If v = 2 Then
  wscript.quit
End if
  r = inputbox("Type in the text that you'd like to have on the Internet Explorer Title bar and then click OK. These input boxes sometimes crash with no text entered  so type 'empty' if you want no text.")
    If r = "empty" Then
          sh.regwrite "HKEY_LOCAL_MACHINE\Software\Microsoft\Internet Explorer\Main\Window Title", "", "REG_SZ"
          sh.regwrite "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main\Window Title", "", "REG_SZ"
     Else
          sh.regwrite "HKEY_LOCAL_MACHINE\Software\Microsoft\Internet Explorer\Main\Window Title", "", "REG_SZ"
          sh.regwrite "HKEY_LOCAL_MACHINE\Software\Microsoft\Internet Explorer\Main\Window Title", r, "REG_SZ"
         sh.regwrite "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main\Window Title", "", "REG_SZ"
         sh.regwrite "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main\Window Title", r, "REG_SZ"
     End if
       msgbox "If you open Internet Explorer now you should see the text you typed at the top.", 0, "Finished"














