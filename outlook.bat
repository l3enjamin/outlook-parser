@echo off
REM Outlook Bridge - Windows entry point using uv
REM This runs on Windows and uses uv to manage dependencies

setlocal

set SCRIPT_DIR=%~dp0
set "MAPPED_DRIVE="

REM Try to map UNC path to drive letter using pushd
pushd "%SCRIPT_DIR%" >nul 2>&1
if %ERRORLEVEL% equ 0 (
    set SCRIPT_DIR=%CD%
    set "MAPPED_DRIVE=1"
)

cd /d "%SCRIPT_DIR%"

REM Use uv to run with pywin32 dependency (module execution)
uv run --with pywin32 -m mailtool %*

REM Popd if we pushed a UNC path
if defined MAPPED_DRIVE (
    popd
)
