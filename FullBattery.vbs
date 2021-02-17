set oLocator = CreateObject("WbemScripting.SWbemLocator")
set oServices = oLocator.ConnectServer(".","root\wmi")
set oResults = oServices.ExecQuery("select * from batteryfullchargedcapacity")
for each oResult in oResults
   iFull = oResult.FullChargedCapacity
next

while (1)
  set oResults = oServices.ExecQuery("select * from batterystatus")
  for each oResult in oResults
    iRemaining = oResult.RemainingCapacity
    bCharging = oResult.Charging
  next
  iPercent = ((iRemaining / iFull) * 100) mod 100
  if bCharging and (iPercent > 92) Then 
    msgbox "Battery is at " & iPercent & "%!! Unplug The Charger!",vbInformation, "Notifikasi Baterai - Battery Notification"
  elseif bCharging and (iPercent < 36) Then 
  elseif (Not bCharging) and (iPercent < 36) Then
    msgbox "Battery is low at " & iPercent & "%!! Plug In The Charger!",vbCritical, "Notifikasi Baterai - Battery Notification"
  end if
  wscript.sleep 30000 ' 5 minutes
wend