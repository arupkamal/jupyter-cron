Set run = WScript.CreateObject("WScript.Shell")
NoteBookPath = "C:\Users\kamala1\Python"
run.Run "jupyter-notebook.exe --no-browser --notebook-dir=" + NoteBookPath, 0, True