Set wshShell = wscript.CreateObject("WScript.Shell")
text = InputBox("Text ?", "Input")
repeat = InputBox("Repeat ?", "Input")
delay = InputBox("Delay ?", "Input")
wscript.sleep 2000
For i = 0 To repeat
wshshell.sendkeys text
wshshell.sendkeys "{ENTER}"
wscript.sleep delay
Next