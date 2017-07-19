Set wshNetwork = CreateObject("WScript.Network")

Const map_srv2_j  ="cn=map_srv2_j"
Const map_srv1_k  ="cn=map_srv1_k"
Const map_srv3_w  ="cn=map_srv3_w"


Set ADSysInfo = CreateObject("ADSystemInfo") 
Set CurrentUser = GetObject("LDAP://" & ADSysInfo.UserName)

if TypeName(CurrentUser.MemberOf) = "String" Then
     strGroups = LCase(CurrentUser.MemberOf)
else
     strGroups = LCase(Join(CurrentUser.MemberOf))
End if

On Error Resume Next


If InStr(strGroups, map_srv2_j) Then
    wshNetwork.MapNetworkDrive "j:", "\\srv2\app\appl" 
End if

If InStr(strGroups, map_srv1_k) Then
	wshNetwork.MapNetworkDrive "k:", "\\srv1\home" 
End if

If InStr(strGroups, map_srv3_w) Then
	wshNetwork.MapNetworkDrive "w:", "\\srv3\project" 
End if

dim POPUP_WAIT, Shell, Env, sDate, h, strMsg, GreetingTime 
POPUP_WAIT = 40
Set Shell = WScript.CreateObject("WScript.Shell")
set Env = Shell.Environment("PROCESS")
'Soobshenie
sDate = Now
h = Hour(sDate)

if (h < 12)  then 
   GreetingTime = "Good morning! "
  elseif (h < 17) then
   GreetingTime = "Good day "
else       
   GreetingTime = "Good evening! "
end if

strMsg = GreetingTime & VbCrLf & LastName & "Welcome to domain." & VbCrLf _
  & "Now  " & FormatDateTime(now(), VbLongTime) & ".  " & GetTodayLongDate

Shell.Popup strMsg, POPUP_WAIT, "Login Script", 0
wscript.quit
