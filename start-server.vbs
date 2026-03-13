Set WshShell = CreateObject("WScript.Shell")
WshShell.Run "cmd /c cd /d C:\Users\txp190010\my-webpage && python -m http.server 9000", 0, False
WshShell.Run "cmd /c cd /d C:\Users\txp190010\my-webpage\pdf-cleaner && python app.py", 0, False
WshShell.Run "cmd /c cd /d C:\Users\txp190010\my-webpage\student-lookup && python app.py", 0, False
WshShell.Run "cmd /c cd /d C:\Users\txp190010\Downloads\AWS-SCS-C03-Study-Docs && python serve.py", 0, False
