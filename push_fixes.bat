@echo off
cd /d C:\Projects\AutoUpdate
git add -A
git commit -m "Fix logger race condition, remove dead ScheduledTaskService, fix output paths"
git push origin main
pause
