Set ws = CreateObject("Wscript.Shell")

ws.Run "powershell.exe -Command ""wsl -d Debian -u admin""", vbhide