Set WshShell = CreateObject("WScript.Shell")
WshShell.CurrentDirectory = "d:\Antigravity\123\telegrambot"
WshShell.Run "python -u sheets_ai_bot.py", 0, False
