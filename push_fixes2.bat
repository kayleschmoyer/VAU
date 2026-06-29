@echo off
cd /d C:\Projects\AutoUpdate
echo === CLEANUP AND PUSH === > C:\Projects\AutoUpdate\push_result2.txt

REM Remove temp files from git tracking and disk
git rm -f cleanup.bat cleanup_result.txt done.txt final_cleanup.bat git_output.txt push_all_fixes.bat push_fix.bat push_result.txt >> C:\Projects\AutoUpdate\push_result2.txt 2>&1

REM Stage the fixed .sln and .vbproj
git add VastAutoUpdater.sln VastAutoUpdater/VastAutoUpdater.vbproj >> C:\Projects\AutoUpdate\push_result2.txt 2>&1

echo === DIFF === >> C:\Projects\AutoUpdate\push_result2.txt
git diff --cached --stat >> C:\Projects\AutoUpdate\push_result2.txt 2>&1

echo === COMMIT === >> C:\Projects\AutoUpdate\push_result2.txt
git commit -m "fix: correct .sln project path and .vbproj backslashes, remove temp files" >> C:\Projects\AutoUpdate\push_result2.txt 2>&1

echo === PUSH === >> C:\Projects\AutoUpdate\push_result2.txt
git push origin main >> C:\Projects\AutoUpdate\push_result2.txt 2>&1

echo === DONE === >> C:\Projects\AutoUpdate\push_result2.txt
