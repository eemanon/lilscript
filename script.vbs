Do
Dim WShell
Set WShell = CreateObject("WScript.Shell")
WShell.Run "my.hta", 1, True
Set WShell = Nothing
WScript.Sleep 10000
Loop