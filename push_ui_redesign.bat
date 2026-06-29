@echo off
cd /d C:\Projects\AutoUpdate
echo === UI REDESIGN PUSH === > C:\Projects\AutoUpdate\push_ui_result.txt

REM Remove old Form1 files from tracking
git rm -f VastAutoUpdater/UI/Form1.vb VastAutoUpdater/UI/Form1.Designer.vb VastAutoUpdater/UI/Form1.resx >> C:\Projects\AutoUpdate\push_ui_result.txt 2>&1

REM Remove leftover temp files
git rm -f push_fixes2.bat push_result2.txt 2>> C:\Projects\AutoUpdate\push_ui_result.txt

REM Stage new and modified files
git add VastAutoUpdater/UI/MainForm.vb VastAutoUpdater/UI/MainForm.Designer.vb VastAutoUpdater/UI/MainForm.resx >> C:\Projects\AutoUpdate\push_ui_result.txt 2>&1
git add VastAutoUpdater/VastAutoUpdater.vbproj >> C:\Projects\AutoUpdate\push_ui_result.txt 2>&1
git add "VastAutoUpdater/My Project/Application.Designer.vb" >> C:\Projects\AutoUpdate\push_ui_result.txt 2>&1
git add VastAutoUpdater/Services/UpdaterEngine.vb >> C:\Projects\AutoUpdate\push_ui_result.txt 2>&1

echo === DIFF === >> C:\Projects\AutoUpdate\push_ui_result.txt
git diff --cached --stat >> C:\Projects\AutoUpdate\push_ui_result.txt 2>&1

echo === COMMIT === >> C:\Projects\AutoUpdate\push_ui_result.txt
git commit -m "feat: redesign UI with Klipboard branding" -m "- Rename VASTUpdater/Form1 to MainForm" -m "- Drop MaterialSkin dependency for form, use custom-painted WinForms" -m "- Gradient magenta header with grid pattern overlay" -m "- Brand colors: Magenta #ED017F, Charcoal #333333, Biscuit White #F2F3EF" -m "- Inter font throughout, modern flat controls" -m "- Borderless window with drag support" -m "- Magenta progress bar via Win32 PBM_SETBARCOLOR" -m "- Fix Await in Finally block (BC36943) in UpdaterEngine.vb" >> C:\Projects\AutoUpdate\push_ui_result.txt 2>&1

echo === PUSH === >> C:\Projects\AutoUpdate\push_ui_result.txt
git push origin main >> C:\Projects\AutoUpdate\push_ui_result.txt 2>&1

echo === DONE === >> C:\Projects\AutoUpdate\push_ui_result.txt
