@echo off
cd /d C:\Projects\AutoUpdate
git rm -f push_fix.bat push_all_fixes.bat cleanup.bat git_output.txt cleanup_result.txt push_result.txt final_cleanup.bat
git commit -m "chore: remove temp build scripts"
git push origin main
echo CLEANUP DONE > C:\Projects\AutoUpdate\done.txt
