@echo off
cd /d "%~dp0"
where git >nul 2>nul || (echo Git belum terpasang. Pasang dari https://git-scm.com lalu jalankan file ini lagi. & pause & exit /b)
if not exist ".git" git init -q -b main
git config user.email >nul 2>nul || (git config user.name "ajojing-dotcom" & git config user.email "ajojing-dotcom@users.noreply.github.com")
git remote get-url origin >nul 2>nul || git remote add origin https://github.com/ajojing-dotcom/pindai.git
git add -A
git commit -q -m "Perbarui Pindai" >nul 2>nul
echo Mengirim ke GitHub. Bila diminta, login ke akun ajojing-dotcom di jendela yang muncul.
git push -u origin main
if errorlevel 1 (echo. & echo Push gagal. Kirim pesan error di atas ke Claude. & pause & exit /b)
echo.
echo Selesai: https://github.com/ajojing-dotcom/pindai
pause
