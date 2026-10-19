@echo off
where python >nul 2>nul || (echo Python belum terpasang. Pasang Python dari python.org, centang "Add python.exe to PATH", lalu jalankan file ini lagi. & pause & exit /b)
python -c "import win32com.client, PIL, twain" 2>nul || python -m pip install pywin32 pillow pytwain
python "%~dp0pindai.py" --shortcut
start "" pythonw "%~dp0pindai.py"
