var WshShell = new ActiveXObject("WScript.Shell");
WScript.sleep(2000);
for (var i = 0; i < 20; i++) {
	WshShell.SendKeys("");
	WshShell.SendKeys("{ENTER}");
	WScript.sleep(100);
}