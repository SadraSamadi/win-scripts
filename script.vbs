Set WshShell = WScript.CreateObject("WScript.Shell")
WScript.Sleep 2000
Do While true
WshShell.SendKeys Time
WshShell.SendKeys "{ENTER}"
WScript.Sleep 1000
Loop
