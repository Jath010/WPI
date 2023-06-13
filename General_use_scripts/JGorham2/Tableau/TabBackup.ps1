Rename-Item c:\backups\prdtab06.tsbak prdtab07.tsbak
Rename-Item c:\backups\prdtab05.tsbak prdtab06.tsbak
Rename-Item c:\backups\prdtab04.tsbak prdtab05.tsbak
Rename-Item c:\backups\prdtab03.tsbak prdtab04.tsbak
Rename-Item c:\backups\prdtab02.tsbak prdtab03.tsbak
Rename-Item c:\backups\prdtab01.tsbak prdtab02.tsbak
Rename-Item c:\backups\prdtab.tsbak prdtab01.tsbak
& "tsm" maintenance backup --file prdtab.tsbak
Move-Item "C:\ProgramData\Tableau\Tableau Server\data\tabsvc\files\backups\prdtab.tsbak" C:\backups -Force
Copy-Item c:\backups\prdtab.tsbak \\tsttableau-01\c$\backups\prdtab.tsbak
Remove-Item c:\backups\prdtab07