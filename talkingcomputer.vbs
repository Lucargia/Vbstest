Dim msg, sapi
msg=InputBox("Enter you text","Talk it")
Set sapi=CreateObject("Sapi.spvoice")
sapi.speak msg
