@echo off
cd /d C:\Projects\AutoUpdate
git add -A
git commit -m "Code review fixes: template credentials, path traversal guard, atomic download, exit code check, font fallback, assembly info, gradient guard, FormClosing handler"
git push origin main
pause
