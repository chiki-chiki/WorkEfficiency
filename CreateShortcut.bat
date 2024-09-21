@echo off

REM ショートカットのパス
set "shortcut=C:\Users\t-usami\Desktop\ショートカット\test.lnk"

REM 新しいリンク先のパス
set "newTarget=C:\Program Files (x86)\Magicxpa\Studio 3.2\MgxpaStudio.exe"

REM ショートカットのリンク先を書き換える
powershell -Command "(New-Object -ComObject WScript.Shell).CreateShortcut('%shortcut%').TargetPath = '%newTarget%'; (New-Object -ComObject WScript.Shell).CreateShortcut('%shortcut%').Save()"
