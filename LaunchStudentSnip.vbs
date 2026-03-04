Option Explicit

Dim fso, shell, baseDir, pywPath, appPath, cmd, exitCode, q
Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")
q = Chr(34)

baseDir = fso.GetParentFolderName(WScript.ScriptFullName)
pywPath = baseDir & "\.venv\Scripts\pythonw.exe"
appPath = baseDir & "\app.py"

If fso.FileExists(pywPath) And fso.FileExists(appPath) Then
    shell.CurrentDirectory = baseDir
    cmd = q & pywPath & q & " " & q & appPath & q
    exitCode = shell.Run(cmd, 0, True)
    If exitCode <> 0 Then
        MsgBox "StudentSnip exited with code " & exitCode & "." & vbCrLf & _
               "Try running app.py manually to see the error.", vbExclamation, "StudentSnip Launcher"
    End If
Else
    MsgBox "Could not find .venv\Scripts\pythonw.exe or app.py." & vbCrLf & _
           "Set up the .venv and dependencies first.", vbExclamation, "StudentSnip Launcher"
End If
