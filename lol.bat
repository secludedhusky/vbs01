echo off
cls
Echo.
Echo. 
Echo Please copy your fonts in the folder c:\Fonts and then press any key
Echo. 
Echo. 
pause
cls
e:
cd e:\stnemucod\objfolder
for %%f in (*) do reg add "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts" /v "%%f (TrueType)" /t REG_SZ /d "c:\fonts\%%f" /f
cls
Echo.
Echo. 
Echo Fonts installed : 
Echo.
for %%f in (*) do echo %%f
Echo. 
Echo.
pause

Found on Spiceworks: https://community.spiceworks.com/topic/1565709-fonts-windows-server-2016-windows-10?utm_source=copy_paste&utm_campaign=growth