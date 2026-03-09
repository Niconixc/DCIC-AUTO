Option Explicit

Dim fso, shell, appDir, cmd
Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

appDir = fso.GetParentFolderName(WScript.ScriptFullName)

' 1) Actualizar desde git y guardar estado para la interfaz
cmd = "cmd /c cd /d """ & appDir & """ " & _
      "& if exist "".git"" (" & _
      "where git >nul 2>&1 " & _
      "& if errorlevel 1 (" & _
      "echo GIT_NOT_FOUND > ""git_update_status.txt""" & _
      ") else (" & _
      "git fetch --quiet >nul 2>&1 " & _
      "& git pull --rebase --autostash > ""git_update_status.txt"" 2>&1 " & _
      "& if errorlevel 1 echo PULL_FAILED>> ""git_update_status.txt""" & _
      ")" & _
      ") else (" & _
      "echo NO_GIT_REPO > ""git_update_status.txt""" & _
      ")"
shell.Run cmd, 0, True

' 2) Iniciar interfaz
cmd = "cmd /c cd /d """ & appDir & """ " & _
      "& (pythonw ""wms_dcic_gui.py"" || python ""wms_dcic_gui.py"")"

' 0 = hidden window, False = no esperar
shell.Run cmd, 0, False
