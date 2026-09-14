@echo off
title AutoInterview - IUO Letter and WhatsApp Dispatcher
color 0A

:: 1. Prioritize Standalone Executable (Zero installation required)
if exist "%~dp0AutoInterview.exe" (
    "%~dp0AutoInterview.exe" %*
    goto :check_error
)

:: 2. Fallback to Python if source repository without binary
set PYTHON_EXE=""
if exist "%USERPROFILE%\miniconda3\python.exe" (
    set PYTHON_EXE="%USERPROFILE%\miniconda3\python.exe"
) else if exist "%LOCALAPPDATA%\miniconda3\python.exe" (
    set PYTHON_EXE="%LOCALAPPDATA%\miniconda3\python.exe"
) else (
    set PYTHON_EXE=python
)

%PYTHON_EXE% auto_interview.py %*

:check_error
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERROR] An error occurred during execution.
    pause
)
