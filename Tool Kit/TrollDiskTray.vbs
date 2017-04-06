@echo off 
title CD Drive Prank
echo Press Any Key When Ready To Activate
pause>
Set oWMP=CreateObject("WMPlayer.ocx.7")
Set colCDROMs=oWMP.cdromCollection
do if colCDROMs.Count>=1 then 
For i =0 to colCDROMs.Count -1
colCDROMs.Item(i).Eject
next
For i =0 to colCDROMs.Count -1
colCDROMs.Item(i).Eject
next 
End If
wscript.sleep 100
loop