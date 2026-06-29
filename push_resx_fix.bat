@echo off
cd /d C:\Projects\AutoUpdate
echo === RESX FIX PUSH === > C:\Projects\AutoUpdate\push_resx_result.txt

git add VastAutoUpdater/UI/MainForm.resx >> C:\Projects\AutoUpdate\push_resx_result.txt 2>&1

echo === DIFF === >> C:\Projects\AutoUpdate\push_resx_result.txt
git diff --cached --stat >> C:\Projects\AutoUpdate\push_resx_result.txt 2>&1

echo === COMMIT === >> C:\Projects\AutoUpdate\push_resx_result.txt
git commit -m "fix: repair truncated MainForm.resx closing tags" >> C:\Projects\AutoUpdate\push_resx_result.txt 2>&1

echo === PUSH === >> C:\Projects\AutoUpdate\push_resx_result.txt
git push origin main >> C:\Projects\AutoUpdate\push_resx_result.txt 2>&1

echo === DONE === >> C:\Projects\AutoUpdate\push_resx_result.txt
