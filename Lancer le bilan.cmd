@echo off
rem Pactio - Bilan express : double-cliquez sur ce fichier pour lancer le bilan.
start "" powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "%~dp0Bilan-Express_Lanceur.ps1"
