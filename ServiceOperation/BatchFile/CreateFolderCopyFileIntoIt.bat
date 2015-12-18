REM #### Get the current time for the log entry ####
For /f "tokens=1-2 delims=/:" %%a in ('time /t') do (set Currtime=%%a%%b)

FOR /F "TOKENS=1 DELIMS=/ " %%A IN ('DATE /T') DO SET dd=%%A
FOR /F "TOKENS=2 DELIMS=/ " %%A IN ('DATE /T') DO SET mm=%%A
FOR /F "TOKENS=3 DELIMS=/ " %%A IN ('DATE /T') DO SET yyyy=%%A
set Pdate=%yyyy%%mm%%dd%
REM #### Map the drive and provide UNC credentials ####
net use Z: \\10.0.0.246\AAProducts testuser /USER:ALIGNEDASSETS\testuser
md "Z:\AutomationReport\BLUELIGHT_NAG_IX\IMPORT\ORACLE_OS%Pdate%%Currtime%\"
md "Z:\AutomationReport\BLUELIGHT_NAG_IX\IMPORT\SQL_OS%Pdate%%Currtime%\"
md "Z:\AutomationReport\BLUELIGHT_NAG_IX\IMPORT\ORACLE_LO%Pdate%%Currtime%\"
md "Z:\AutomationReport\BLUELIGHT_NAG_IX\IMPORT\SQL_LO%Pdate%%Currtime%\"
md "Z:\AutomationReport\BLUELIGHT_NAG_IX\IMPORT\ORACLE_FRC%Pdate%%Currtime%\"
md "Z:\AutomationReport\BLUELIGHT_NAG_IX\IMPORT\SQL_FRC%Pdate%%Currtime%\"
md "Z:\AutomationReport\BLUELIGHT_NAG_IX\IMPORT\ORACLE_LOFRC%Pdate%%Currtime%\"
md "Z:\AutomationReport\BLUELIGHT_NAG_IX\IMPORT\SQL_LOFRC%Pdate%%Currtime%\"

copy /V /Y "C:\Automation\BL_iEx\CSV\Import Files\COU_Report\ORACLE\*.*" "Z:\AutomationReport\BLUELIGHT_NAG_IX\IMPORT\ORACLE_OS%Pdate%%Currtime%\"
copy /V /Y "C:\Automation\BL_iEx\CSV\Import Files\COU_Report\SQL\*.*" "Z:\AutomationReport\BLUELIGHT_NAG_IX\IMPORT\SQL_OS%Pdate%%Currtime%\"
copy /V /Y "C:\Automation\BL_iEx\CSV\Import Files\LO_Report\ORACLE\*.*" "Z:\AutomationReport\BLUELIGHT_NAG_IX\IMPORT\ORACLE_LO%Pdate%%Currtime%\"
copy /V /Y "C:\Automation\BL_iEx\CSV\Import Files\LO_Report\SQL\*.*" "Z:\AutomationReport\BLUELIGHT_NAG_IX\IMPORT\SQL_LO%Pdate%%Currtime%\"
copy /V /Y "C:\Automation\BL_iEx\CSV\Import Files\Add_Rep_FRC\ORACLE\*.*" "Z:\AutomationReport\BLUELIGHT_NAG_IX\IMPORT\ORACLE_FRC%Pdate%%Currtime%\"
copy /V /Y "C:\Automation\BL_iEx\CSV\Import Files\Add_Rep_FRC\SQL\*.*" "Z:\AutomationReport\BLUELIGHT_NAG_IX\IMPORT\SQL_FRC%Pdate%%Currtime%\"
copy /V /Y "C:\Automation\BL_iEx\CSV\Import Files\Add_Rep_LO\ORACLE\*.*" "Z:\AutomationReport\BLUELIGHT_NAG_IX\IMPORT\ORACLE_LOFRC%Pdate%%Currtime%\"
copy /V /Y "C:\Automation\BL_iEx\CSV\Import Files\Add_Rep_LO\SQL\*.*" "Z:\AutomationReport\BLUELIGHT_NAG_IX\IMPORT\SQL_LOFRC%Pdate%%Currtime%\"




Rem copy /V /Y "C:\Automation\SinglePoint\AddressBase\WebSite\Database\TestPlanAfterExecution\Oracle_Environment\*.xls" "Z:\SinglePoint\AutomationTestPlan\Automation\TestReport\ReportOn%Pdate%%Currtime%\"