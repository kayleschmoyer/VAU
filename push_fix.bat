@echo off
cd /d C:\Projects\AutoUpdate
git add -A
git commit -m "fix: silent mode reads SFTP credentials from ConfigManager"
git push origin main
echo.
echo === DONE ===
pause
