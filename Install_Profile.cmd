@ECHO OFF

set PWD=%~dp0
echo *** Setting Execution Policy to Unrestricted
powershell.exe -command "& {set-executionpolicy unrestricted -ea 1}"
if %ERRORLEVEL% NEQ 0 (goto error) else (echo  -- Done.)

echo *** Unblocking files in repository
powershell.exe -ExecutionPolicy ByPass -Command "&{import-module '%PWD%\Modules\Security\Security.psm1' -Force -ea 1; dir '%PWD%' -rec | Unblock-File -Verbose}"
if %ERRORLEVEL% NEQ 0 (goto error) else (echo  -- Done.)

echo *** Calling profile installation PowerShell script
powershell.exe -ExecutionPolicy ByPass -File "%~dp0\Install_Profile.ps1"
if %ERRORLEVEL% NEQ 0 (goto error) else (echo  -- Done.)

goto eof

:error
echo "Failed to install profile. The last command didn't succeed..."
set /p "Press Enter to Continue"

:eof