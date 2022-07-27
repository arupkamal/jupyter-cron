NoteBookPath = "C:\Users\kamala1\Python"

ComputerName = "."
Set objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\cimv2")
sQuery = "SELECT * FROM Win32_Process WHERE Name LIKE '%jupyter-notebook%'"
Set objItems = objWMIService.ExecQuery(sQuery)
res = objItems.count

IF res = 0 THEN 
    Set run = WScript.CreateObject("WScript.Shell")
    run.Run "jupyter-notebook.exe --no-browser --notebook-dir=" + NoteBookPath, 0, True
ELSE 
    WScript.Echo "Jupyter-Notebook is already running"
END IF
