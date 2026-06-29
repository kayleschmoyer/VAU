@echo off
cd /d C:\Projects\AutoUpdate
echo === GIT STATUS === > C:\Projects\AutoUpdate\push_result.txt
git status >> C:\Projects\AutoUpdate\push_result.txt 2>&1
echo === GIT ADD === >> C:\Projects\AutoUpdate\push_result.txt
git add -A >> C:\Projects\AutoUpdate\push_result.txt 2>&1
echo === GIT DIFF STAGED === >> C:\Projects\AutoUpdate\push_result.txt
git diff --cached --stat >> C:\Projects\AutoUpdate\push_result.txt 2>&1
echo === GIT COMMIT === >> C:\Projects\AutoUpdate\push_result.txt
git commit -m "fix: comprehensive hardening - logging, error handling, security, retry logic" -m "- Logger: fallback file log at ProgramData\VASTUpdater\Logs\ when Event Log unavailable" -m "- ApplicationEvents: global UnhandledException handler prevents silent crashes" -m "- ScheduledTaskService: call schtasks.exe directly instead of via cmd.exe (injection fix)" -m "- UpdaterEngine: retry logic (3 attempts), safe progress calc, installer verification" -m "- SftpService: connection/operation timeouts, proper Using-block disposal" -m "- ConfigManager: TryParse safety, DPAPI encryption support for credentials" -m "- Form1: hide UI in silent mode, centralized exit, safe Invoke calls" -m "- VersionService: cache discovered path, skip non-fixed drives" -m "- InstallerPathService: deduplicated folder creation, old installer cleanup" -m "- EmailService: config validation, machine name in emails, SMTP timeout" -m "- VastAutoUpdater.vbproj: Option Strict On" >> C:\Projects\AutoUpdate\push_result.txt 2>&1
echo === GIT PUSH === >> C:\Projects\AutoUpdate\push_result.txt
git push origin main >> C:\Projects\AutoUpdate\push_result.txt 2>&1
echo === DONE === >> C:\Projects\AutoUpdate\push_result.txt
