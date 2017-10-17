@ECHO OFF

if exist L:\ (
    ECHO Removing existing L:\...
    Net use L: /delete
)

ECHO Mapping eduHub share to L:\...
Net use L: \\10.0.0.1\eduHub$ Password /User:DOMAIN\eduhubUser

ECHO Running Import Script...
PowerShell.exe -ExecutionPolicy Bypass -Command "& '%~dpn0.ps1' 'L:\OnDemand.csv'" >> L:\OnDemand.log

ECHO Removing eduHub share...
Net use L: /delete

ECHO Done.
timeout /t 5