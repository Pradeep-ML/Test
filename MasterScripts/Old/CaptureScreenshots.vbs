Set objWSH = CreateObject("WScript.Shell")
For i=1 To 10
	objWSH.SendKeys "^%{F12}"
	WScript.Sleep 500
Next