@echo off
cd /d C:\Projects\AutoUpdate
git add -A
git commit -m "Fix BC36943 Await-in-Catch, fix XML doc comment placement warnings"
git push origin main
pause
