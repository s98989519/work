@echo off

setlocal enabledelayedexpansion

set file=%~dsp0
start /wait %file%MECM_action_script_all.bat
intl.cpl
