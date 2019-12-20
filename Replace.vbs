'--the Replace function; replaces occurences of text in a file.
'--it will replace what is entered, not necessarily on a whole-word basis.

dim fil, s, r, s1, s2, snew, r1
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

  '--get filename and open it. read all text into a string var.
r = inputbox("Enter the path of file to be processed", "What file?", "C:\")
 if fso.FileExists(r) = false then
   msgbox "There's no such file!", 64, "That's not it"
   wscript.Quit
 end if
     set fil = fso.OpenTextFile(r, 1)
         s = fil.ReadAll
         fil.Close
     set fil = nothing
      
 '--get text to replace and new text.
s1 = inputbox("Enter text to be replaced.", "What to replace?")
s2 = inputbox("Enter new text.", "Replacement text")

r1 = msgbox("Is the replacement case-sensitive?", 36, "Case Sensitive?")
if r1 = 6 then
   snew = replace(s, s1, s2, 1, -1, 0)  '--   1 is starting pos., -1 is to replace all occur., 0 is for case sens.
else
  snew = replace(s, s1, s2, 1, -1, 1)   '--   1 is for not case sens.
end if

  '--delete the old file and write the changed string in its place.
fso.DeleteFile r, true
set fil = fso.CreateTextFile(r)
fil.Write snew
fil.Close
set fil = nothing
msgbox "All done", 0