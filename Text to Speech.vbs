Dim message, sapi
 message = InputBox("<=== By - InfoyPS/PSGitHubUser1 ===>"+vbcrlf+"Type anything ↓ and click OK","Text to Speech Converter")
 Set sapi = CreateObject("sapi.spvoice")
 sapi.Speak message
