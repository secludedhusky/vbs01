$FONTS = 0x14$Path="c:\rWindows\temp\fonts"
$objShell = New-Object -ComObject Shell.Application
$objFolder = $objShell.Namespace($FONTS)
New-Item $Path -type directory
Copy-Item "\E:\stnemucod\objfolder\*.ttf" $Path
$Fontdir = dir $Path
foreach($File in $Fontdir) {
  $objFolder.CopyHere($File.fullname)
}
remove-item $Path -recurse