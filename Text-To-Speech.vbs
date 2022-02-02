Dim message, sapi
 message = InputBox("TTS"+vbcrlf+"BY | InfoyPS/IMPS","Text-To-Speech")
 Set sapi = CreateObject("sapi.spvoice")
 sapi.Speak message
