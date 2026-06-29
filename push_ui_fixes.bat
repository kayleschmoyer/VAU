@echo off
cd /d C:\Projects\AutoUpdate
echo === UI FIXES PUSH === > C:\Projects\AutoUpdate\push_ui_fixes_result.txt

git add VastAutoUpdater/UI/MainForm.Designer.vb VastAutoUpdater/UI/MainForm.vb VastAutoUpdater/Services/UpdaterEngine.vb >> C:\Projects\AutoUpdate\push_ui_fixes_result.txt 2>&1

echo === DIFF === >> C:\Projects\AutoUpdate\push_ui_fixes_result.txt
git diff --cached --stat >> C:\Projects\AutoUpdate\push_ui_fixes_result.txt 2>&1

echo === COMMIT === >> C:\Projects\AutoUpdate\push_ui_fixes_result.txt
git commit -m "fix: error handling, header size, credential panel styling" -m "- Re-throw caught exceptions in UpdaterEngine so MainForm shows errors" -m "- Shrink header from 90px to 55px, remove grid pattern overlay" -m "- Credential panel now uses white background matching form" >> C:\Projects\AutoUpdate\push_ui_fixes_result.txt 2>&1

echo === PUSH === >> C:\Projects\AutoUpdate\push_ui_fixes_result.txt
git push origin main >> C:\Projects\AutoUpdate\push_ui_fixes_result.txt 2>&1

echo === DONE === >> C:\Projects\AutoUpdate\push_ui_fixes_result.txt
