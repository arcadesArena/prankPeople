option explicit
dim sp,oShell,user,comp,text
Set oShell = CreateObject( "WScript.Shell" )
user=oShell.ExpandEnvironmentStrings("%UserName%")
comp=oShell.ExpandEnvironmentStrings("%ComputerName%")
text="text "
msgbox comp
set sp=createobject("sapi.spvoice")
do
'sp.speak text & user
msgbox(text & user)
wscript.sleep 1000
loop