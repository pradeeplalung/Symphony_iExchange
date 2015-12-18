REM #### Get the current time for the log entry ####
For /f "tokens=1-2 delims=/:" %%a in ('time /t') do (set Currtime=%%a%%b)

FOR /F "TOKENS=1 DELIMS=/ " %%A IN ('DATE /T') DO SET dd=%%A
FOR /F "TOKENS=2 DELIMS=/ " %%A IN ('DATE /T') DO SET mm=%%A
FOR /F "TOKENS=3 DELIMS=/ " %%A IN ('DATE /T') DO SET yyyy=%%A
set Pdate=%yyyy%%mm%%dd%

REM #### Map the drive and provide UNC credentials ####
REM net use R: \\10.0.0.246\project_data testuser /USER:ALIGNEDASSETS\testuser
REM net use Z: \\10.0.0.246\AAProducts testuser /USER:ALIGNEDASSETS\testuser

copy /V /Y "R:\Symphony Bluelight iExchange\Testing\Automation\AddressBase\BatchFile\BX_TP13\ORACLE\*.bat" "C:\Automation\BL_iEx\BatchFile\BX_TP13\ORACLE\"

