option explicit
dim sp
set sp=createobject("sapi.spvoice")
do
sp.speak "text"
wscript.sleep 1000
loop