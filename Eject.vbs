Set objCdDrive = CreateObject("WMPlayer.OCX.7")
Set colCDROM = objCdDrive.cdromCollection
do
MsgBox "Ha Ha Ha Ha ..."
colCDROM.Item(0).Eject
wscript.sleep 1000
colCDROM.Item(0).Eject
wscript.sleep 10000
Loop