@echo off
rem ダブルクリックやタスクスケジューラから実行するためのランチャー
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0ArrivalPlanMerger.ps1"
if not "%1"=="/nopause" pause
