@echo off
title Tally Import Generator

set LOCKFILE=app.lock

:: Get current PID
for /f %%i in ('powershell -command "$PID"') do set MYPID=%%i

:: If lock file exists
if exist %LOCKFILE% (
    set /p OLDPID=<%LOCKFILE%

    tasklist /FI "PID eq %OLDPID%" 2>NUL | find "%OLDPID%" >NUL
    if not errorlevel 1 (
        powershell -command "[System.Windows.Forms.MessageBox]::Show('Tally Import Generator is already running.','Application Already Running',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)"
        exit /b
    ) else (
        echo Removing stale lock...
        del %LOCKFILE%
    )
)

:: Write current PID into lock file
echo %MYPID% > %LOCKFILE%

:: Run app
python launcher.py

:: Remove lock file after exit
if exist %LOCKFILE% del %LOCKFILE%

exit
