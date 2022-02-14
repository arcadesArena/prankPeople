option explicit
dim sp,txtmsg
set sp=createobject("sapi.spvoice")
txtmsg="msg that is converted to speech"
do
sp.speak txtmsg
wscript.sleep 1000
loop