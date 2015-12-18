REM #### Get the current time for the log entry ####
For /f "tokens=1-2 delims=/:" %%a in ('time /t') do (set Currtime=%%a%%b)

FOR /F "TOKENS=1 DELIMS=/ " %%A IN ('DATE /T') DO SET dd=%%A
FOR /F "TOKENS=2 DELIMS=/ " %%A IN ('DATE /T') DO SET mm=%%A
FOR /F "TOKENS=3 DELIMS=/ " %%A IN ('DATE /T') DO SET yyyy=%%A
set Pdate=%yyyy%%mm%%dd%

REM #### Map the drive and provide UNC credentials ####
REM net use Z: \\10.0.0.246\AAProducts testuser /USER:ALIGNEDASSETS\testuser

Rem md "C:\SinglePoint\AutomationTestPlan\Automation\AddressBase\LogFile\LogFileSQLOn%Pdate%%Currtime%"

REM copy /V /Y "Z:\Symphony\Bluelight iExchange\Automation\AddressBase\BatchFile\Bluelight_iExchange_OracleAutomation.bat" "C:\Automation\BL_iEx\BatchFile\"
REM copy /V /Y "Z:\Symphony\Bluelight iExchange\Automation\AddressBase\BatchFile\Bluelight_iExchange_SQLAutomation.bat" "C:\Automation\BL_iEx\BatchFile\"
REM copy /V /Y "Z:\Symphony\Bluelight iExchange\Automation\AddressBase\BatchFile\T001_OS_CreateFolderCopyFileIntoIt.bat" "C:\Automation\BL_iEx\BatchFile\"
REM copy /V /Y "Z:\Symphony\Bluelight iExchange\Automation\AddressBase\BatchFile\T002_LO_CreateFolderCopyFileIntoIt.bat" "C:\Automation\BL_iEx\BatchFile\"
REM copy /V /Y "Z:\Symphony\Bluelight iExchange\Automation\AddressBase\BatchFile\T003_FRC_CreateFolderCopyFileIntoIt.bat" "C:\Automation\BL_iEx\BatchFile\"
REM copy /V /Y "Z:\Symphony\Bluelight iExchange\Automation\AddressBase\BatchFile\T004_LO_FRC_CreateFolderCopyFileIntoIt.bat" "C:\Automation\BL_iEx\BatchFile\"
copy /V /Y "Z:\Symphony\Bluelight iExchange\Automation\AddressBase\BatchFile\*.bat" "C:\Automation\BL_iEx\BatchFile\"