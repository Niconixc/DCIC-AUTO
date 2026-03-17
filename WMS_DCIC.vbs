Option Explicit

Dim fso, shell, appDir, cmd
Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

appDir = fso.GetParentFolderName(WScript.ScriptFullName)

' Iniciar la interfaz (Python maneja el git pull internamente al arrancar)
cmd = "cmd /c cd /d """ & appDir & """ & (pythonw ""wms_dcic_gui.py"" || python ""wms_dcic_gui.py"")"

' 0 = ventana oculta, False = no esperar
shell.Run cmd, 0, False
