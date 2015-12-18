Rem  Save the application log list entry from the event viewer
powershell -ExecutionPolicy Bypass -NoLogo -NoProfile -Command "Get-EventLog -Log "Application" | Export-CSV C:\Automation\BL_iEx\ServiceOperation\SaveEventLogfile\ORA_Environment\ServiceLog.csv"
