@echo off
cd /d C:\Projects\AutoUpdate
git rm push_fix.bat git_output.txt cleanup.bat
git commit -m "chore: remove temp files"
git push origin main
echo CLEANUP DONE > C:\Projects\AutoUpdate\cleanup_result.txt
