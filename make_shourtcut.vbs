set WshShell = WScript.CreateObject("WScript.Shell")
desktop = WshShell.SpecialFolders("Desktop")

set objShortcut = WshShell.CreateShortcut(desktop & "\ToDo List.lnk")
''実行ファイルを設定
objShortcut.TargetPath = WshShell.CurrentDirectory & "\ToDoList.exe"
''アイコンを設定
objShortcut.IconLocation = WshShell.CurrentDirectory & "\ToDo.ico"
''ショートカットを保存
objShortcut.Save