REM #### Get the current time for the log entry ####
For /f "tokens=1-2 delims=/:" %%a in ('time /t') do (set Currtime=%%a%%b)

FOR /F "TOKENS=1 DELIMS=/ " %%A IN ('DATE /T') DO SET dd=%%A
FOR /F "TOKENS=2 DELIMS=/ " %%A IN ('DATE /T') DO SET mm=%%A
FOR /F "TOKENS=3 DELIMS=/ " %%A IN ('DATE /T') DO SET yyyy=%%A
set Pdate=%yyyy%%mm%%dd%

REM #### Map the drive and provide UNC credentials ####
net use Z: \\10.0.0.246\AAProducts testuser /USER:ALIGNEDASSETS\testuser

copy /V /Y "Z:\Symphony\Bluelight iExchange\Automation\AddressBase\TestPlan_SQL\*.xls" "C:\Automation\BL_iEx\Database\"
copy /V /Y "Z:\Symphony\Bluelight iExchange\Automation\AddressBase\CSVFile\COU\*.csv" "C:\Automation\BL_iEx\CSV\Import Files\COU\"
copy /V /Y "Z:\Symphony\Bluelight iExchange\Automation\AddressBase\CSVFile\LO\*.csv" "C:\Automation\BL_iEx\CSV\Import Files\LO\"
copy /V /Y "Z:\Symphony\Bluelight iExchange\Automation\AddressBase\CSVfile\Additional\*.csv" "C:\Automation\BL_iEx\CSV\Import Files\Additional\"



