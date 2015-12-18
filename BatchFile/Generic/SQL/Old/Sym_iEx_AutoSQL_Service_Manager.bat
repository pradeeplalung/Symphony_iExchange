REM #### Get the current time for the log entry ####
For /f "tokens=1-2 delims=/:" %%a in ('time /t') do (set Currtime=%%a:%%b)

FOR /F "TOKENS=1 DELIMS=/ " %%A IN ('DATE /T') DO SET dd=%%A
FOR /F "TOKENS=2 DELIMS=/ " %%A IN ('DATE /T') DO SET mm=%%A
FOR /F "TOKENS=3 DELIMS=/ " %%A IN ('DATE /T') DO SET yyyy=%%A
set Pdate=%yyyy%_%mm%_%dd%


REM #### Set the logfile name and location ####
set SymLog="C:\Automation\Sym_iEx\AutomatedInstall\InstallLog\AutoProcess%PDate%.log"

REM #### Start the LOG for this process ####
echo Job started at (Year, Month, Day, time) : %Pdate% %Currtime% >> %SymLog%

REM #### Map the drive and provide UNC credentials ####
REM net use Z: \\10.0.0.246\AAProducts testuser /USER:ALIGNEDASSETS\testuser
if errorlevel 1 goto FAIL_PART1

REM #### UNC drive mapped successfully ####
echo Mapped AAProducts drive  >> %SymLog%

REM ####    COPY FILE FROM QTP FOLDER   ####

REM #### Copy files from QTP Folder to local Folder, if it fails STOP everything else as there's no point continuing. ####
copy /V /Y "Z:\Symphony\Bluelight iExchange\Setup\Automation\*.exe" "C:\Automation\BL_iEx\AutomatedInstall\"
REM copy /V /Y "Z:\Symphony\Bluelight\Setup\Automation\Symphony Bluelight SQL Server 5.5*.exe" "C:\Automation\BL_iEx\AutomatedInstall\"if errorlevel 1 goto FAIL_PART2

REM #### The file was found and copied successfully ####
echo Executable Files Moved Successfully >> %SymLog%

REM #### Assign files to Variables ####
for %%A IN ("C:\Automation\BL_iEx\AutomatedInstall\Bluelight iExchange 5*.*") DO set Bluelight="%%A"
for %%B IN ("C:\Automation\BL_iEx\AutomatedInstall\Bluelight iExchange Manager*.*") DO set Manager="%%B"
REM for %%D IN ("C:\Automation\BL_iEx\AutomatedInstall\Bluelight iExchange SQL*.*") DO set SQL="%%D"
REM for %%D IN ("C:\Bluelight\AutomatedInstall\Symphony Bluelight SQL*.*") DO set SQL="%%D"




REM #### Save FileName  in Log File ####
echo "###### Files to Install ########" >>%SymLog%
echo %Bluelight% >> %SymLog%
echo %Manager% >> %SymLog%
echo %SQL% >> %SymLog%
echo "################################" >>%SymLog%

REM #### Attempt to run files ####
%Bluelight% SILENT=TRUE TARGETDIR="C:\Program Files\Aligned Assets Limited\Bluelight iExchange\" ALLUSERS=True Name=Alignedassets Company=alignedassets Serial1=4519 Serial2=5267071 Serial3=7331 Serial4=3325082 ORACLE_11G=FALSE ORACLE_10G=FALSE SQL_SERVER=TRUE /l="C:\Automation\BL_iEx\AutomatedInstall\InstallLog\Bluelight_iExchange_Log.txt"
if errorlevel 1 goto FAIL_PART3

REM #### Attempt to run files ####
%Manager% SILENT=TRUE TARGETDIR="C:\Program Files\Aligned Assets Limited\Bluelight iExchange\" ALLUSERS=True Name=Alignedassets Company=alignedassets Serial1=4519 Serial2=5267071 Serial3=7331 Serial4=3325082 ORACLE_11G=FALSE ORACLE_10G=FALSE SQL_SERVER=TRUE /l="C:\Automation\BL_iEx\AutomatedInstall\InstallLog\Manager_Log.txt"
if errorlevel 1 echo (Part3) Manager failed to install, Not essential item >> %SymLog%

REM ***************Added the below code ***
REM #### Attempt to run files ####
REM %SQL% SILENT=TRUE 1  SERVER_NAME="PL-Win7Dual" DATABASE_NAME="BLUELIGHT_NAG_RECIP" /l="C:\Bluelight\AutomatedInstall\InstallLog\SQL_Log.txt"
REM if errorlevel 1 goto FAIL_PART4


REM #### After Successful processing Skip the failed Messages, each fail message MUST end in "goto :END" in order to skip other messages. ####
REM schtasks /change /tn "TaskNameGoesHere" /DISABLE
REM set InstLog="C:\Automation\BL_iEx\AutomatedInstall\InstallLog\InstallLogSQL\Install_Successful_%Pdate%.log"
REM echo Install Succeeded : >> %InstLog%
REM goto :END

REM ************End*****

REM #### After Successful processing Skip the failed Messages, each fail message MUST end in "goto :END" in order to skip other messages. ###
REM schtasks /change /tn "TaskNameGoesHere" /DISABLE
set InstLog="C:\Automation\BL_iEx\AutomatedInstall\InstallLog\InstallLogSQLService\Install_Successful_%Pdate%.log"
echo Install Succeeded : >> %InstLog%
goto :END

:FAIL_PART1
REM #### File copy failed so nothing was run ####
echo (Part1) Failed to map unc drive, process aborted >> %SymLog%
goto :END

:FAIL_PART2
REM #### File copy failed so nothing was run ####
echo (Part2) Executable files not found and so failed to copy, process aborted >> %SymLog%
goto :END

:FAIL_PART3
REM #### File install failed so exit ####
echo (Part3) Bluelight iExchange failed to install, process aborted >> %SymLog%
goto :END

:FAIL_PART4
REM #### File install failed so exit ####
echo (Part4) SQL Database Migration failed to install, process aborted >> %SymLog%
goto :END

:END
REM #### Get the current time for the log entry ####
For /f "tokens=1-2 delims=/:" %%a in ('time /t') do (set Currtime=%%a:%%b)
echo Job Ended at (Year, Month, Day, time) : %Pdate% %Currtime% >> %SymLog%

rem #### Disconnect the mapped drive ####
REM net use Z: /delete
