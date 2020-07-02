Set WshShell = WScript.CreateObject("WScript.Shell")

WScript.Sleep 5000

Do While True

WshShell.SendKeys "{CAPSLOCK}"
WScript.Sleep 2000
WshShell.SendKeys "{CAPSLOCK}"
WScript.Sleep 5000

Loop
