REM #### Get the current time for the log entry ####
For /f "tokens=1-2 delims=/:" %%a in ('time /t') do (set Currtime=%%a%%b)

FOR /F "TOKENS=1 DELIMS=/ " %%A IN ('DATE /T') DO SET dd=%%A
FOR /F "TOKENS=2 DELIMS=/ " %%A IN ('DATE /T') DO SET mm=%%A
FOR /F "TOKENS=3 DELIMS=/ " %%A IN ('DATE /T') DO SET yyyy=%%A
set Pdate=%yyyy%%mm%%dd%


REM md "Z:\Symphony\Bluelight iExchange\Automation\AddressBase\CSVReport_Recipient\ORA\ORA_OS%Pdate%%Currtime%\"
REM **** The following code - Copy the entire folders within "C:\TEMP\System_Repository" in network drive ****
REM XCOPY /E /I "C:\TEMP\System_Repository" "Z:\Symphony\Bluelight iExchange\Automation\AddressBase\CSVReport_Recipient\ORA\ORA_OS%Pdate%%Currtime%\"
copy /V /Y "C:\Automation\BL_iEx\Database\BX_TP01_NSG_Core.xls" "R:\Symphony Bluelight iExchange\Testing\Automation\AddressBase\TestPlan_ORA\"
