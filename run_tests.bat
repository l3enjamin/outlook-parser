@echo off
REM Test Runner Script for Windows
REM Uses uv to run pytest with pywin32 dependency

setlocal

REM Get the directory where this script is located
set SCRIPT_DIR=%~dp0
set "MAPPED_DRIVE="

REM Try to map UNC path to drive letter using pushd
pushd "%SCRIPT_DIR%" >nul 2>&1
if %ERRORLEVEL% equ 0 (
    set SCRIPT_DIR=%CD%
    set "MAPPED_DRIVE=1"
)

REM Change to script directory
cd /d "%SCRIPT_DIR%"

REM Ensure .venv is valid - recreate if broken
if exist .venv (
    echo Checking .venv...
    where .venv\Scripts\python.exe >nul 2>&1
    if errorlevel 1 (
        where .venv\bin\python >nul 2>&1
        if errorlevel 1 (
            echo Broken .venv detected, recreating...
            uv venv --recreate >nul 2>&1
        )
    )
)

REM Default arguments if none provided
if "%~1"=="" (
    set PYTEST_ARGS=-v
) else (
    set PYTEST_ARGS=%*
)

echo ========================================
echo Mailtool Outlook Bridge Test Suite
echo ========================================
echo.
echo Running pytest with arguments: %PYTEST_ARGS%
echo Current directory: %CD%
echo.

REM Run pytest using uv with pywin32
REM Use --no-project to skip broken project venv
REM Add tests directory to PYTHONPATH for conftest imports
set PYTHONPATH=%CD%\tests;%PYTHONPATH%
uv run --no-project --with pywin32 --with pytest --with pytest-timeout pytest %PYTEST_ARGS%

REM Capture exit code
set EXIT_CODE=%ERRORLEVEL%

echo.
echo ========================================
if %EXIT_CODE% equ 0 (
    echo Tests PASSED
) else (
    echo Tests FAILED ^(exit code: %EXIT_CODE%^)
)
echo ========================================

REM Popd if we pushed a UNC path
if defined MAPPED_DRIVE (
    popd
)

exit /b %EXIT_CODE%
