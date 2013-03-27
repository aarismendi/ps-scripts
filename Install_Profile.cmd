set PWD=%~dp0
powershell.exe -executionpolicy bypass -command "&{import-module '%PWD%\Modules\Security\Security.psm1' -Force; dir '%PWD%' -rec | Unblock-File -Verbose}"
powershell.exe -ExecutionPolicy ByPass -File "%~dp0\Install_Profile.ps1"