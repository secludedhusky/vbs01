
'--Script to make shortcut links in IE favorites
'--*************************************************
Dim fso, ok, LnkURL, LnkName, LnkPath, subfol, fol, sh, Lnk, Favfol

Set fso = CreateObject("Scripting.FileSystemObject")
Set sh = CreateObject("WScript.Shell")

ok = false
   
  '--get link target:
LnkURL = inputbox("Enter the URL of the website you want to add to the Favorites menu in IE, like 'http://www.hotbot.com'.", "Web Address")
  If LnkURL = "" then
      wscript.quit
  end if

  '--get link filename:
LnkName = inputbox("Enter the name you want to see for the link, like 'Hotbot'.", "Link Name")
   If LnkName = "" then
    wscript.quit
  end if
 '--get favorites folder:
  Favfol = sh.Specialfolders("Favorites")

  '--option to put link in a subfolder of Favorites:
subfol = inputbox("If you want the link in an existing favorites folder then enter the name of the folder here.", "Link Location")
If not subfol = "" then
        do until ok
          If not subfol = "" then 
            If not fso.folderexists(Favfol & "\" & subfol) then
              subfol = inputbox("That folder doesn't exist. You can rewrite it or leave it blank to simply put it in Favorites.", "Woops!")
            Else
              ok = true
            end if
          else
              ok = true
          end if
        loop
 end if

  If not subfol = "" then
    fol = favfol & "\" & subfol
  else
    fol = favfol
  end if 

   r = msgbox("A link will be made named " & LnkName & ", pointing to " & LnkURL & vbcrlf & "Continue?", 36)
   If r = 7 then 
     wscript.quit
   else    
          '--create the link:
     Set Lnk = sh.CreateShortcut(fol & "\" & LnkName & ".URL")
     Lnk.TargetPath = LnkURL
     Lnk.Save
     msgbox "The link is made.", 0, "All Set"
   end if