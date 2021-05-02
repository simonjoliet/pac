@ECHO off

REM Setting the myPass and myMach var
SET myPass=0
SET myMach=0

REM Setting putty's registry's path
SET puttyRegKey=HKEY_CURRENT_USER\Software\SimonTatham\PuTTY\Sessions\Default%%20settings

REM Setting the configuration file directory
SET confFile=%~dp0\conf.ini

REM If no argument where given, exiting
IF [%1]==[] exit /b
IF [%2]==[] exit /b

REM Getting the password of the machine
ini.vbs %1 %2 "%confFile%" //NOLOGO > tmp.txt
SET /p myPass=<tmp.txt
REM if the password wasn't retreived, exiting
IF [%myPass%]==[0] exit /b

REM Getting the network name of the remote machine
ini.vbs %1 name "%confFile%" //NOLOGO > tmp.txt
SET /p myMach=<tmp.txt
REM If machine's name wasn't retreived, exiting
IF [%myMach%]==[0] exit /b

REM Getting the background color of the specific host
ini.vbs %1 bgcolor "%confFile%" //NOLOGO > tmp.txt
SET /p myBgcolor=<tmp.txt

REM Getting the font color of the specific host
ini.vbs %1 ftcolor "%confFile%" //NOLOGO > tmp.txt
SET /p myFtcolor=<tmp.txt

REM Then deleting the temp file
DEL tmp.txt

REM Setting the specific background and font color in the registry
REG ADD %puttyRegKey% /V Colour0 /T REG_SZ /D %myFtcolor% /F
REG ADD %puttyRegKey% /V Colour2 /T REG_SZ /D %myBgcolor% /F

REM Starting Putty with the retrieved name/password
START PUTTY -ssh %2@%myMach%:22 -pw %myPass%

REM Then exiting CMD
EXIT /b