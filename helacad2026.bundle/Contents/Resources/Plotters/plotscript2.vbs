Set objShell = CreateObject("WScript.Shell")
wscript.sleep (500)
objshell.sendkeys "{ESC}"
objshell.SendKeys "op "
wscript.sleep (500)
objshell.sendkeys "{ESC}"
objshell.SendKeys "plot "
wscript.Quit