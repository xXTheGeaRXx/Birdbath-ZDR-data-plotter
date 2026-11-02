' Parancsikon létrehozása a radar_plotter.py-hoz
' Futtasd ezt egyszer, hogy az Asztalon megjelenjen a parancsikon.

Dim WShell, Shortcut, ScriptDir
Set WShell = CreateObject("WScript.Shell")
ScriptDir = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)

Shortcut = WShell.CreateShortcut(WShell.SpecialFolders("Desktop") & "\Radar Plotter.lnk")
Shortcut.TargetPath = ScriptDir & "\Inditas.bat"
Shortcut.WorkingDirectory = ScriptDir
Shortcut.WindowStyle = 1
Shortcut.Description = "Radar adatok vizualizátora"
Shortcut.Save

MsgBox "Parancsikon sikeresen létrehozva az Asztalon!", vbInformation, "Radar Plotter"
