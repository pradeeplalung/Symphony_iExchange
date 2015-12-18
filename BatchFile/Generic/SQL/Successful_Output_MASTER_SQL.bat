REM #### Get the current time for the log entry ####
For /f "tokens=1-2 delims=/:" %%a in ('time /t') do (set Currtime=%%a%%b)

FOR /F "TOKENS=1 DELIMS=/ " %%A IN ('DATE /T') DO SET dd=%%A
FOR /F "TOKENS=2 DELIMS=/ " %%A IN ('DATE /T') DO SET mm=%%A
FOR /F "TOKENS=3 DELIMS=/ " %%A IN ('DATE /T') DO SET yyyy=%%A
set Pdate=%yyyy%%mm%%dd%

copy /V /Y "C:\Automation\Sym_iEx\Database\IX_MASTERBASELINE_SQL_READY.xls" "R:\Symphony iExchange\Testing\Automation\iManage\LLPG\Run_Completion_MASTER_SQL\"