REM #### Map the drive and provide UNC credentials ####
REM net use Z: \\10.0.0.246\AAProducts testuser /USER:ALIGNEDASSETS\testuser

REM #### Get the current time for the log entry ####
For /f "tokens=1-2 delims=/:" %%a in ('time /t') do (set Currtime=%%a%%b)

FOR /F "TOKENS=1 DELIMS=/ " %%A IN ('DATE /T') DO SET dd=%%A
FOR /F "TOKENS=2 DELIMS=/ " %%A IN ('DATE /T') DO SET mm=%%A
FOR /F "TOKENS=3 DELIMS=/ " %%A IN ('DATE /T') DO SET yyyy=%%A
set Pdate=%yyyy%%mm%%dd%
md "R:\Symphony iExchange\Testing\Automation\iManage\LLPG\TestResults\IX_TP11\SQL\T001\ImportReport\IR_%Pdate%%Currtime%\"

copy /V /Y "C:\Automation\Sym_iEx\CSV\Import Files\NLPG_Report\SQL\*.*" "R:\Symphony iExchange\Testing\Automation\iManage\LLPG\TestResults\IX_TP11\SQL\T001\ImportReport\IR_%Pdate%%Currtime%\"