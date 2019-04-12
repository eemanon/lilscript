Dim WShell
Do
Set WShell = CreateObject("WScript.Shell")
WScript.Sleep 10000
WShell.Run "my.hta", 1, True
Set WShell = Nothing
Loop