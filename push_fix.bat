@echo off
cd /d C:\Projects\AutoUpdate
echo === GIT STATUS === > C:\Projects\AutoUpdate\git_output.txt
git status >> C:\Projects\AutoUpdate\git_output.txt 2>&1
echo === GIT ADD === >> C:\Projects\AutoUpdate\git_output.txt
git add -A >> C:\Projects\AutoUpdate\git_output.txt 2>&1
echo === GIT COMMIT === >> C:\Projects\AutoUpdate\git_output.txt
git commit -m "fix: silent mode reads SFTP credentials from ConfigManager" >> C:\Projects\AutoUpdate\git_output.txt 2>&1
echo === GIT PUSH === >> C:\Projects\AutoUpdate\git_output.txt
git push origin main >> C:\Projects\AutoUpdate\git_output.txt 2>&1
echo === DONE === >> C:\Projects\AutoUpdate\git_output.txt
