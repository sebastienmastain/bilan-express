@echo off
rem Pactio - Bilan express : double-cliquez sur ce fichier pour lancer le bilan.
rem Ouvre la fenetre des options, puis le bilan dans une fenetre PowerShell visible.
rem Lecture seule : rien n'est modifie sur le PC.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Bilan-Express_Lanceur.ps1"
