Set WshShell = CreateObject("WScript.Shell")
WshShell.Run "cmd /c start_windows.bat > start_windows.log 2>&1", 0, false