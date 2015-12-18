REM #### Get the current time for the log entry ####
REM For /f "tokens=1-2 delims=/:" %%a in ('time /t') do (set Currtime=%%a%%b)

REM FOR /F "TOKENS=1 DELIMS=/ " %%A IN ('DATE /T') DO SET dd=%%A
REM FOR /F "TOKENS=2 DELIMS=/ " %%A IN ('DATE /T') DO SET mm=%%A
REM FOR /F "TOKENS=3 DELIMS=/ " %%A IN ('DATE /T') DO SET yyyy=%%A
REM set Pdate=%yyyy%%mm%%dd%
net use Z: \\10.0.0.246\AAProducts testuser /USER:ALIGNEDASSETS\testuser
DEL Z:\SinglePoint\AutomationTestPlan\Automation\AddressBase\Run_Completion_ORA\*.xls