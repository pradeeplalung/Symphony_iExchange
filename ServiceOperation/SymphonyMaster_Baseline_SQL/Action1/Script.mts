''Test Objectives : Copy and pasting required batch files and xls files from network drive
Option Explicit
Print "********************************************************"
Dim vStartTime: vStartTime = Time
Dim vMsgServiceStarted,vMsgServiceStopped,vServiceName,vSourceName,vComputerName,vStartedSuccessfully,vStoppedSuccessfully
Dim objWMIService,colLoggedEvents,colServiceList,objService,objEvent
Dim vOracle_Environment ,vSQL_Environment
Dim vServiceStarted,vServiceStopped
Const  vStarted = "Service Started"
Const  vStopped= "Service Stopped"

Dim vApp_Path,vTP_Path,vTP,vTCD,vTCD1,vTCD2,vTCD3,vTCD4,vTSEC,vTPS,vAppStatus

Dim vInstallLog,vInstallLogPathORA,vInstallLogPathSQL,vInstallOra,vInstallSql,vInstallPath,vExtension,vLogName

Dim vWaitInstallORALog,vWaitInstallSQLLog,vWaitInstallFilePathORA,vWaitInstallFilePathSQL,vWaitFileSQL,vWaitFileORA 

Dim vMigrationCheck_ORA,vMigrationCheck_SQL,vIM_DBMigration_SQL,vIM_DBMigration_SQL_Status,vSinglePointTest,vIX_DBMigration_SQL

Dim vSinglePointTestONOFF,vDBMigrationONOFF,vIX_DBMigrationONOFF,vIX_DBMigration_SQL_Status,vSinglePointInstallation

Dim vOraDBUpgrade,vSqlDBUpgrade,vExistControllerORA,vExistControllerSQL

Dim vSym_iX_Service,vSym_iX_Service1,vSym_iX_Service_Date,vSym_iX_Version,vSym_iX_Version1,vSym_iX_Manager,vSym_iX_Manager1

Dim vSym_iX_Manager_Version,vSym_iX_Manager_Version1,vSym_iX_Manager_Date,vSym_DB_Build,vSym_DB_Build1,vSym_DB_Version

Dim vSym_DB_Version1,vSym_DB_Build_Date,vAutoProcess_LogFile,vSym_iX_Ver_Paths,vCsvFile_ORA,vCsvFile_SQL

Dim vSym_IX_DBPath_Len_SQL,vSym_IX_Mng_BldPath_Len,vSym_BLD_BldPath_Len_ORA,vSym_BLD_BldPath_Len_SQL,vSym_IX_BldPath_Len

Dim vSym_IXDb,vSym_IXDb1,vSym_iX_DB_Date,vSym_IXDb_Ver,vSym_IXDb_Ver1

Dim vSP_Build,vSP_Ver_Paths,vSP_Build_Date,vSP_Build1,vSP_Version,vSP_Version1,vSystem_Folder

Dim vTo,vCC,vSym_AllBlds_Folder,vAutoProcess_Log,vSym_IX_LogTxt_Path,vImportReportFiles_ORA,vImportReportFiles_SQL

'Dim vInitializationVBScript_ORA, vBaseLineMaster_ORA,vReplaceScriptName_ORA
Dim vInitializationVBScript_SQL,vSymphonyMaster_Baseline_SQL,vReplaceScriptName_SQL

Dim vWaitInstallFilePathSQLiX,vWaitInstallFilePathORAService,vExistControllerSQLSinglePoint,vWaitInstallFilePathSQLSinglePoint
Dim vSP_BldPath_Len,vExistControllerSQLIX

'Dim vTestEnvironment_ORA:vTestEnvironment_ORA = "ORA"
Dim vTestEnvironment_SQL:vTestEnvironment_SQL = "SQL"

vDBMigrationONOFF = Null
vIX_DBMigrationONOFF = Null
vSinglePointTestONOFF = Null
vIX_DBMigration_SQL_Status = Null
vSinglePointInstallation = Null
vSym_DB_Version1 = "Released DB Version"
vSym_DB_Build_Date = "Released DB Version Date"

'
'				****** STEP1 ***************
'				@ Open network drive and 
'				@ Copy test plans, CSV files, xml file and batch file to QTP/UFT Machine
				Systemutil.Run "C:\Automation\Sym_iEx\BatchFile\Generic\SQL\NWDCon_AppDB_SQL.bat"
				Wait(15)
'				****** STEP2 ***************				
'''				@ Assigned  Environment variable  values dynamically from XML file *************
				Environment.LoadFromFile("C:\Automation\Sym_iEx\Xml_File\Sym_TP_Data_Generic.xml")

				vMsgServiceStarted = Environment.Value("vMsgServiceStarted")
				print "vMsgServiceStarted  :"&vMsgServiceStarted

				vMsgServiceStopped = Environment.Value("vMsgServiceStopped")

				vServiceName = Environment.Value("vServiceName")

				vSourceName = Environment.Value("vSourceName")

				vComputerName = Environment.Value("vComputerName")

				vInstallLogPathORA = Environment.Value("vInstallLogPathORA")

				vInstallLogPathSQL = Environment.Value("vInstallLogPathSQL")
				

				vInstallOra = Environment.Value("vInstallOra")
		
				vInstallSql = Environment.Value("vInstallSql")
				
				vInstallPath = Environment.Value("vInstallPath")

				vExtension = Environment.Value("vExtension")
'
				vLogName = Environment.Value("vLogName")

				
				vApp_Path = Environment.Value("vApp_Path")
				
				vTP_Path = Environment.Value("vTP_Path")
				print "vTP_Path  :"&vTP_Path
'				
				vTPS = Environment.Value("vTPS")
				print "vTPS :"&vTPS
'			
				vTCD1 = Environment.Value("vTCD1")
'				
				vTCD2 = Environment.Value("vTCD2")
'				
				vTCD3 = Environment.Value("vTCD3")
'				
				vTCD4 = Environment.Value("vTCD4")
				
'				
				vTSEC = Environment.Value("vTSEC")
				
				vWaitInstallFilePathSQL = Environment.Value("vWaitInstallFilePathSQL")
				vWaitInstallFilePathSQLiX = Environment.Value("vWaitInstallFilePathSQLiX")
				print "vWaitInstallFilePathSQLiX :"&vWaitInstallFilePathSQLiX
'				vWaitInstallFilePathORAService = Environment.Value("vWaitInstallFilePathORAService")
				
		

				vInstallLog = Environment.Value("vInstallLog")
				
				vInstallPath = Environment.Value("vInstallPath")
				vSym_IX_LogTxt_Path = Environment.Value("vSym_IX_LogTxt_Path")
				
				print "vSym_IX_LogTxt_Path  :"&vSym_IX_LogTxt_Path
				vAutoProcess_Log = Environment.Value("vAutoProcess_Log")
				
				print "vAutoProcess_Log  :"&vAutoProcess_Log
				
				vDefaultLenghtBldPath = Environment.Value("vDefaultLenghtBldPath")
				
				vCsvFile_ORA = Environment.Value("vCsvFile_ORA")
				vCsvFile_SQL = Environment.Value("vCsvFile_SQL")
				
				vImportReportFiles_ORA = Environment.Value("vImportReportFiles_ORA")
				print "vImportReportFiles_ORA  :"&vImportReportFiles_ORA
				vImportReportFiles_SQL= Environment.Value("vImportReportFiles_SQL")
				print "vImportReportFiles_SQL  :"&vImportReportFiles_SQL
				
				vSym_IX_BldPath_Len = Environment.Value("vSym_IX_BldPath_Len")
				vSym_IX_Mng_BldPath_Len = Environment.Value("vSym_IX_Mng_BldPath_Len")
				vSym_IX_DBPath_Len_SQL = Environment.Value("vSym_IX_DBPath_Len_SQL")
				
				vSym_BLD_BldPath_Len_ORA = Environment.Value("vSym_BLD_BldPath_Len_ORA")
				vSym_BLD_BldPath_Len_SQL = Environment.Value("vSym_BLD_BldPath_Len_SQL")
								
				vSP_BldPath_Len = Environment.Value("vSP_BldPath_Len")
				
'				vInitializationVBScript_ORA = Environment.Value("vInitializationVBScript_ORA")
				vInitializationVBScript_SQL = Environment.Value("vInitializationVBScript_SQL")
'				vBaseLineMaster_ORA = environment.Value("vBaseLineMaster_ORA")
				vSymphonyMaster_Baseline_SQL = environment.Value("vSymphonyMaster_Baseline_SQL")
'				vReplaceScriptName_ORA = Environment.Value("vReplaceScriptName_ORA")
				vReplaceScriptName_SQL = Environment.Value("vReplaceScriptName_SQL")
				
				vWaitInstallFilePathSQLSinglePoint = Environment.Value("vWaitInstallFilePathSQLSinglePoint")
				
			    print "vWaitInstallFilePathSQLSinglePoint :"&vWaitInstallFilePathSQLSinglePoint
			    
'			    ************* Setting Symphony iExchange ******
				vSystem_Folder = Environment.Value("vSystem_Folder")
				print "vSystem_Folder  :"&vSystem_Folder
'				****** STEP3 ***************	
'				@ Read the information from Test Plan
				vTo = fnReadFromExcel(vTP_Path,vTPS,12,2)
				print "vTo  :"&vTo
				vCC = fnReadFromExcel(vTP_Path,vTPS,13,2)
				print "vCC  :"&vCC
				
				vSym_AllBlds_Folder = Trim(fnReadFromExcel(vTP_Path,vTPS,9,2))

				vSym_DB_Build = Trim(fnReadFromExcel(vTP_Path,vTPS,19,2)) ' "C:\Automation\Sym_iEx\AutomatedInstall\Symphony Symphony Oracle 5.5.0.0.0.0.exe"			
		

				vIM_DBMigration_SQL= fnReadFromExcel(vTP_Path,vTPS,16,2)
				print "vIM_DBMigration_SQL  :"&vIM_DBMigration_SQL 
				
				vIX_DBMigration_SQL= fnReadFromExcel(vTP_Path,vTPS,24,2)
				print "IX_DBMigration_SQL  :"&vIX_DBMigration_SQL

'				**********************************************************************************
								
				vSinglePointTest = fnReadFromExcel(vTP_Path,vTPS,27,2)
				print "vSinglePointTest  :"&vSinglePointTest
				
				
				
			
 
'********************* STEP4 for Database Migration SQL ***************	
If vIM_DBMigration_SQL = "Yes" Then
			print "Hello Kela1"
			vIM_DBMigration_SQL_Status = fnIM_DBMigration_SQL
''           vIM_DBMigration_SQL_Status = True

End If
'**************
If vIM_DBMigration_SQL = "Yes" and vIM_DBMigration_SQL_Status = True and vIX_DBMigration_SQL = "Yes" Then
			print "Hello Kela2"
			vDBMigrationONOFF = "ON"
			vIX_DBMigrationONOFF = "ON"
			vIX_DBMigration_SQL_Status = fnIX_DBMigrationIns_SQL()
'			**** Setting and saving Symphony iExchange ****
			fniExSetting vSystem_Folder
'			*******************
			systemutil.Run "C:\Automation\Sym_iEx\BatchFile\Generic\Sym_iEx_AccSet.bat"
		vIX_DBMigration_SQL_Status = True
End If
'**************
If vIM_DBMigration_SQL = "No" and vIX_DBMigration_SQL = "Yes" Then
			print "Hello Kela3"
			vDBMigrationONOFF = "OFF"
			vIX_DBMigrationONOFF = "ON"
			vIX_DBMigration_SQL_Status = fnIX_DBMigrationIns_SQL()
			'**** Setting and saving Symphony iExchange ****
			fniExSetting vSystem_Folder
'			*******************
			systemutil.Run "C:\Automation\Sym_iEx\BatchFile\Generic\Sym_iEx_AccSet.bat"
''			vIX_DBMigration_SQL_Status = True
End If
'**************

If vSinglePointTest = "Yes" and vIM_DBMigration_SQL = "Yes" and vIM_DBMigration_SQL_Status = True Then
			print "Hello Kela4"
			vDBMigrationONOFF = "ON"
			vSinglePointTestONOFF = "ON"
			vSinglePointInstallation = fnSinglePointInstallation
''			vSinglePointInstallation = True
End If
'**************

If (vSinglePointTest = "Yes" and vIM_DBMigration_SQL = "No" and vIX_DBMigration_SQL = "Yes" and vIX_DBMigration_SQL_Status = True) or (vSinglePointTest = "Yes" and vIM_DBMigration_SQL = "No")  Then
			print "Hello Kela6"
			vIX_DBMigrationONOFF = "ON"
			vDBMigrationONOFF = "OFF"
			vSinglePointTestONOFF = "ON"
			vSinglePointInstallation = fnSinglePointInstallation
'			vSinglePointInstallation = True
End If

'**************

If vSinglePointTest = "Yes" and vIM_DBMigration_SQL = "No" and vIX_DBMigration_SQL = "No" Then
			print "Hello Kela7"
			vIX_DBMigrationONOFF = "OFF"
			vDBMigrationONOFF = "OFF"
			vSinglePointTestONOFF = "ON"
			vSinglePointInstallation = fnSinglePointInstallation
'''			vSinglePointInstallation = True
End If

'**************

If vSinglePointTest = "Yes" Then
			print "Hello Kela8"
			vSinglePointTestONOFF = "ON"
End If
'**************
If vSinglePointTest = "No" Then
			print "Hello Kela9"
			vSinglePointTestONOFF = "OFF"
End If


'********************* STEP 7 - Reporting *************************************
Wait(5)
fnReport_SQL
'*********************
''*********************** Reporting End **************************
'******************** Copy all Batch files into the Baseline Master ****
SystemUtil.Run "C:\Automation\Sym_iEx\BatchFile\Generic\SQL\AllBatch4IX_TestPlan_SQL.bat"

'*******************Update QTPLauncher file to Switch and hand over the control for Symphony iXchange Test Plan after creation of Baseline Master ****

fnReplaceLineWithSpecificValue vInitializationVBScript_SQL,vSymphonyMaster_Baseline_SQL,vReplaceScriptName_SQL,vInitializationVBScript_SQL
'************ FINAL - Copy and paste the Testreport to Network drive
SystemUtil.Run "C:\Automation\Sym_iEx\BatchFile\Generic\SQL\Successful_Output_MASTER_SQL.bat"
	'*******************The End of TEST *************
	
'****All Local function Functions ********
Function fnIM_DBMigration_SQL()
'				@ SQL database migration is completed by the batch file
'				@ The 'fnWaitTillExistsORAfunction' will make the QTP to wait for the migration and provide TRUE once complete
'				@ Once migration completes the batch file 'BL_iEx_AccSet.bat' will set Symphony service operation account
				print "I am starting SQL - Migration"
				Systemutil.Run "C:\Automation\Sym_iEx\BatchFile\Generic\SQL\Sym_AutoSQL_Migration.bat"
	''							msgbox "I am starting SQL"
				vExistControllerSQL  =fnWaitTillExistsSQL(vWaitInstallFilePathSQL)
	'							vExistControllerSQL  =True
				Print "vExistControllerSQL_Only  :"&vExistControllerSQL
				fnIM_DBMigration_SQL = vExistControllerSQL

End Function


Function fnIX_DBMigrationIns_SQL()
'				@ SQL database migration is completed by the batch file
'				@ The 'fnWaitTillExistsORAfunction' will make the QTP to wait for the migration and provide TRUE once complete
'				@ Once migration completes the batch file 'BL_iEx_AccSet.bat' will set Symphony service operation account
				print "I am starting IExchange SQL - Migration"
				Systemutil.Run "C:\Automation\Sym_iEx\BatchFile\Generic\SQL\Sym_iEx_AutoSQL_Migration.bat"
	''							msgbox "I am starting SQL"
				vExistControllerSQLIX  =fnWaitTillExistsSQL(vWaitInstallFilePathSQLiX)
	'							vExistControllerSQL  =True
				Print "vExistControllerSQLIX  :"&vExistControllerSQLIX
				fnIX_DBMigrationIns_SQL = vExistControllerSQLIX

End Function
'************************
Function fnSinglePointInstallation()
				print "Single point installation will be beginning now"
	            SystemUtil.Run "C:\Automation\Sym_iEx\BatchFile\SinglePoint_AA_LLPG_Automation_SQL.bat"	 
				print "vWaitInstallFilePathSQLSinglePoint  :"&vWaitInstallFilePathSQLSinglePoint	             
	            vExistControllerSQLSinglePoint  =fnWaitTillExistsSQL(vWaitInstallFilePathSQLSinglePoint)
				Print "vExistControllerSQLService   :"&vExistControllerSQLSinglePoint 
				fnSinglePointInstallation = vExistControllerSQLSinglePoint
End Function
Function fnReport_SQL()
				
				'*************	STEP 10 Reporting *************	
				'@ Dynamic way to bring installed BUILD's information and put in the report for stakeholders
				'@ It is dependant on log files generated from silent inslallers from Symphony IX service, Manager and DB migration THOUGH
				Dim vEndTime:vEndTime = Time
				Dim vTimeTaken,vExecutionTime,vMigrationCheck
				vTimeTaken = vEndTime - vStartTime
				vExecutionTime = fnExecutionTime(vTimeTaken)
				'**********************
				'Keep list of log files in a file called "C:\Automation\Sym_iEx\AutomatedInstall\LogFileName.txt"
				fnGetListFilesInFolder vInstallLog,vSym_IX_LogTxt_Path 
				
				print "Starting ***********"
				vAutoProcess_LogFile = fnReadOnlyThatLineContainsSpecificText(vSym_IX_LogTxt_Path,vAutoProcess_Log) ' Dynamic AutoProcess file name 
				print "vAutoProcess_LogFile  :"&vAutoProcess_LogFile
				
				vSym_iX_Ver_Paths = vInstallLog&vAutoProcess_LogFile '' Dynamic AutoProcess file path name 
				print "vSym_iX_Ver_Paths  :"&vSym_iX_Ver_Paths
				
				'******************* Symphony iexchange installalion information gathering ****
				
				vSym_iX_Service = fnReadOnlyThatLineContainsSpecificText(vSym_iX_Ver_Paths,"Symphony iExchange 5") ' Search Blue
				
				vSym_iX_Service = Quote&vSym_iX_Service&Quote
				print "vSym_iX_Service  :"&vSym_iX_Service
				
				vSym_iX_Service1 = Trim(mid(vSym_iX_Service,2,vSym_IX_BldPath_Len))
				print "vSym_IX_BldPath_Len  :"&vSym_IX_BldPath_Len
				print "vSym_iX_Service1  :"&vSym_iX_Service1
				
				vSym_iX_Service_Date = fnGetTimeStampOfSpecificFile(vSym_iX_Service1)
				
				vSym_iX_Version =split(vSym_iX_Service1, "\")
				vSym_iX_Version1 = vSym_iX_Version(4)
				print "vSym_iX_Version1  :"&vSym_iX_Version1
				
				'******************* Symphony iexchange Manager installalion information gathering ****
				
				vSym_iX_Manager = fnReadOnlyThatLineContainsSpecificText(vSym_iX_Ver_Paths,"Symphony iExchange Manager")
				print "Hello1"
				
				vSym_iX_Manager = Quote&vSym_iX_Manager&Quote
				print "vSym_iX_Manager  :"&vSym_iX_Manager
				
				vSym_iX_Manager1 = Trim(Mid(vSym_iX_Manager,2,vSym_IX_Mng_BldPath_Len))
				
				vSym_iX_Manager_Date = fnGetTimeStampOfSpecificFile(vSym_iX_Manager1)
				
				vSym_iX_Manager_Version =split(vSym_iX_Manager1, "\")
				vSym_iX_Manager_Version1 = vSym_iX_Version(4)
				
				
'				*********************Symphony iexchange DB installalion information gathering*********
					print "*********New entry - Symphony IExchange DB data *****"
				vSym_IXDb = fnReadOnlyThatLineContainsSpecificText(vSym_iX_Ver_Paths,"iExchange SQL Server")
			
				
				vSym_IXDb = Quote&vSym_IXDb&Quote
				print "vSym_IXDb  :"&vSym_IXDb
				
				vSym_IXDb1 = Trim(Mid(vSym_IXDb,2,vSym_IX_DBPath_Len_SQL))
				Print "vSym_IXDb1   :"&vSym_IXDb1 
				vSym_iX_DB_Date = fnGetTimeStampOfSpecificFile(vSym_IXDb1)
				
				print "vSym_iX_DB_Date  :"&vSym_iX_DB_Date
				
				vSym_IXDb_Ver =split(vSym_IXDb1, "\")
				vSym_IXDb_Ver1 = vSym_IXDb_Ver(4)
				
				print "vSym_IXDb_Ver1  :"&vSym_IXDb_Ver1
				 Print "****************** End 1231332**************"
				
				'******************* Symphony DB migration installalion information gathering ****
				print "Hello2"
				vSym_DB_Build = fnReadOnlyThatLineContainsSpecificText(vSym_iX_Ver_Paths,"Symphony Spatial SQL Server")
				
				vSym_DB_Build = Quote&vSym_DB_Build&Quote
				print "vSym_DB_Build :"&vSym_DB_Build
'				*********************
				'vSym_DB_Build1 = Mid(vSym_DB_Build,1,vDefaultLenghtBldPath)
				vSym_DB_Build1 = Trim(Mid(vSym_DB_Build,2,vSym_BLD_BldPath_Len_SQL)) ' Sql Server 
				
				vSym_DB_Build_Date = fnGetTimeStampOfSpecificFile(vSym_DB_Build1)
				
				print "vSym_DB_Build_Date  :"&vSym_DB_Build_Date
				If vSym_DB_Build_Date = null or vSym_DB_Build_Date = "" Then
					vSym_DB_Build_Date = "Release DB Version Date"
				End If
'				******************
				vSym_DB_Version = split(vSym_DB_Build1, "\")
				vSym_DB_Version1 = vSym_DB_Version(4)

				'******************* SinglePoint installalion information gathering ****
				
				vSP_Build = fnReadOnlyThatLineContainsSpecificText(vSym_iX_Ver_Paths,"SinglePoint")
				print "vSP_Build1  :"&vSP_Build
				
				vSP_Build = Quote&vSP_Build&Quote
				print "vSP_Build2  :"&vSP_Build
				'vSym_DB_Build1 = Mid(vSym_DB_Build,1,vDefaultLenghtBldPath)
				vSP_Build1 = Trim(Mid(vSP_Build,2,vSP_BldPath_Len)) ' Sql Server 
				
				vSP_Build_Date = fnGetTimeStampOfSpecificFile(vSP_Build1)
				print "vSP_Build_Date  :"&vSP_Build_Date

				
				print "vSP_Build_Date  :"&vSP_Build_Date
				print "vSP_Build2a  :"&vSP_Build1
				vSP_Version = split(vSP_Build1, "\")
				vSP_Version1 = vSP_Version(4)
				print "vSP_Version1  :"&vSP_Version1
'				*******************
				
				vMigrationCheck_SQL = fnReadFromExcel(vTP_Path,vTPS,26,2)
				
				vEndTime = Time
				vTimeTaken = vEndTime - vStartTime
				vExecutionTime = fnExecutionTime(vTimeTaken)
				'@ Send email to respective recipients informing about TEST (Build, Build dates, execution time , environment & Test Report repository path(s)
				fnSendReportOnMASTERBaseLineCreation vTo,vCC,vTestEnvironment_SQL,vExecutionTime,vSym_iX_Version1,vSym_iX_Service_Date,vSym_iX_Manager_Version1, _
				vSym_iX_Manager_Date,vSym_DB_Version1,vSym_DB_Build_Date,vMigrationCheck_SQL,vCsvFile_SQL,vImportReportFiles_SQL,vSP_Build_Date,vSP_Version1, _
				vSym_IXDb_Ver1,vSym_iX_DB_Date,vDBMigrationONOFF,vSinglePointTestONOFF,vIX_DBMigrationONOFF
End Function

Function  fnWaitTillExistsStarted(vMsgServiceStarted,vSourceName)  
'''    ' Sleeps until the file exists  
'''    ' The polling interval will increase gradually, but never rises above MAX_WAITTIME  
'''    ' Times out after TIMEOUT msec. Will return false if caused by timeout.  
    Dim waittime, totalwaittime, rep, doAgain,withRepeat,fso,arr(2)
    Const INIT_WAITTIME = 20  
    Const MAX_WAITTIME = 10000  
    Const TIMEOUT = 500000  
    Const SLOPE = 1.1  
	withRepeat = True
    doAgain  = True  
	Set fso = createobject("Scripting.FileSystemObject")
    Do While doAgain  
        waittime = INIT_WAITTIME  
        totalwaittime = 0  
        Do While totalwaittime < TIMEOUT  
            waittime = Int (waittime * SLOPE)  
            If waittime>MAX_WAITTIME Then waittime=MAX_WAITTIME  
            totalwaittime = totalwaittime + waittime  
			Wait(waittime)
''''			**************************
			Set objWMIService = GetObject("winmgmts:" _
				& "{impersonationLevel=impersonate}!\\" & vComputerName & "\root\cimv2")
			
			Set colLoggedEvents = objWMIService.ExecQuery _
				("Select * from Win32_NTLogEvent Where Logfile = 'Application'")
			For Each objEvent in colLoggedEvents

				Print "Message: " & objEvent.Message

				Print "Source Name: " & objEvent.SourceName
				print "*****************"
				If  objEvent.Message =vMsgServiceStarted and  objEvent.SourceName = vSourceName Then
								vServiceStarted = "Service Started"
								print  "Service Started"
								Exit For
								Else
								vServiceStarted = "Service NOT Started"	
								print  "Service NOT Started"								
							end if
			next

'		****************************************
			If  vServiceStarted = vStarted  Then			
					fnWaitTillExistsStarted = True		
					Print "*********Service Started ***********"			
					Exit function
				else
					fnWaitTillExistsStarted = False
					Print "*********Service  NOT Started ***********"	
										
				End If   
        Loop  
        If withRepeat Then  
            rep = Print ("This Service Not Started:" & vbcr & vServiceName & vbcr & vbcr & "Keep trying?", vbRetryCancel+vbExclamation, "Service not Started")  
            doAgain = (rep = vbRetry)  
        Else  
            doAgain = False  
        End If  
    Loop  
   fnWaitTillExistsStarted = False  
End Function
'********************
'***********************
Function  fnWaitTillExistsStopped(vMsgServiceStopped,vSourceName)  
'''    ' Sleeps until the file exists  
'''    ' The polling interval will increase gradually, but never rises above MAX_WAITTIME  
'''    ' Times out after TIMEOUT msec. Will return false if caused by timeout.  
    Dim waittime, totalwaittime, rep, doAgain,withRepeat,fso,arr(2)
    Const INIT_WAITTIME = 20  
    Const MAX_WAITTIME = 10000  
    Const TIMEOUT = 500000  
    Const SLOPE = 1.1  
	withRepeat = True
    doAgain  = True  
	Set fso = createobject("Scripting.FileSystemObject")
    Do While doAgain  
        waittime = INIT_WAITTIME  
        totalwaittime = 0  
        Do While totalwaittime < TIMEOUT  
            waittime = Int (waittime * SLOPE)  
            If waittime>MAX_WAITTIME Then waittime=MAX_WAITTIME  
            totalwaittime = totalwaittime + waittime  
			Wait(waittime)
''''			**************************
			Set objWMIService = GetObject("winmgmts:" _
				& "{impersonationLevel=impersonate}!\\" & vComputerName & "\root\cimv2")
			
			Set colLoggedEvents = objWMIService.ExecQuery _
				("Select * from Win32_NTLogEvent Where Logfile = 'Application'")
			For Each objEvent in colLoggedEvents

				Print "Message: " & objEvent.Message

				Print "Source Name: " & objEvent.SourceName
				print "*****************"
							If  objEvent.Message =vMsgServiceStopped and  objEvent.SourceName = vSourceName Then
								vServiceStopped = "Service Stopped"
								Exit For
								Else
								vServiceStopped = "Service Running"								
							end if
			next

'		****************************************
			If  vServiceStopped = vStopped  Then			
					fnWaitTillExistsStopped = True		
					Print "*********Stopped Service Stopped ***********"			
					Exit function
				else
					fnWaitTillExistsStopped = False
					Print "*********Stopped Service  Running ***********"	
										
				End If   
        Loop  
        If withRepeat Then  
            rep = Print ("This Service Not Stopped:" & vbcr &vServiceName& vbcr & vbcr & "Keep trying?", vbRetryCancel+vbExclamation, "Service Not Stopped")  
            doAgain = (rep = vbRetry)  
        Else  
            doAgain = False  
        End If  
    Loop  
   fnWaitTillExistsStopped = False  
End Function
'''**********************Startinga Service with service name *******
Function fnStartService(vServiceName,vComputerName)
		Set objWMIService = GetObject("winmgmts:" _
			& "{impersonationLevel=impersonate}!\\" & vComputerName & "\root\cimv2")
		Set colServiceList = objWMIService.ExecQuery _
			("Select * from Win32_Service where Name='"&vServiceName&"'")
		For each objService in colServiceList
			objService.StartService()
		Next
End Function
'''**********************Stopping a Service with service name *******
Function fnStopService(vServiceName,vComputerName)
		Set objWMIService = GetObject("winmgmts:" _
			& "{impersonationLevel=impersonate}!\\" & vComputerName & "\root\cimv2")
		Set colServiceList = objWMIService.ExecQuery _
			("Select * from Win32_Service where Name='"&vServiceName&"'")
		For each objService in colServiceList
			objService.StopService()
		Next
End Function
'*********************** Stopping Services Running Under a Specific Account ******************
Rem - Stops all services running under the hypothetical service account Netsvc.
Function fnStopServiceAccount(vServiceName, vComputerName)
				Set objWMIService = GetObject("winmgmts:" _
					& "{impersonationLevel=impersonate}!\\" &  vComputerName & "\root\cimv2")
				Set colServices = objWMIService.ExecQuery _
					("Select * from win32_Service where Name='"&vServiceName&"'")
				For each objService in colServices 
					If objService.StartName = ".\alignedassets" Then
						errReturnCode = objService.StopService()
					End If
				Next
End Function
'***************************
Function fnImportBuild()
End Function

'C:\Automation\ServiceOperation\SaveEventLogfiles



''''***************
Function fnWaitTillExistsORA(vWaitInstallFilePathORA)  
'''    ' Sleeps until the file exists  
'''    ' The polling interval will increase gradually, but never rises above MAX_WAITTIME  
'''    ' Times out after TIMEOUT msec. Will return false if caused by timeout.  
    Dim waittime, totalwaittime, rep, doAgain,withRepeat,fso
    Const INIT_WAITTIME = 20  
    Const MAX_WAITTIME = 1000  
    Const TIMEOUT = 50000  
    Const SLOPE = 1.1  
	withRepeat = True
''''    file = replace (file, Chr(34), "") 'remove double quotes from the input  
    doAgain  = True  
''''    Set WshShell = CreateObject( "WScript.Shell") 
	Set fso = createobject("Scripting.FileSystemObject")
    Do While doAgain  
        waittime = INIT_WAITTIME  
        totalwaittime = 0  
        Do While totalwaittime < TIMEOUT  
            waittime = Int (waittime * SLOPE)  
            If waittime>MAX_WAITTIME Then waittime=MAX_WAITTIME  
            totalwaittime = totalwaittime + waittime  
            Wait(waittime)
'''			'**************************
'			vWaitInstallFilePathORA = vInstallPath&vInstallOra
			vWaitInstallORALog = fnGetASingleFileInFolder(vWaitInstallFilePathORA)
			vWaitFileORA = vWaitInstallFilePathORA&"\"&vWaitInstallORALog 
			'	REM ******************* Verifying Migration Check  warning dialog box **********
            If fso.fileExists (vWaitFileORA) Then  
					fnWaitTillExistsORA = True  
					Exit Function 
				else  
						If  fnMigrationCheck = "Migration Check" Then
								fnWriteToExcel vTP_Path,vTPS,25,2,"Yes"
								Print "Yes"
								else 
'								fnWriteToExcel vTP_Path,vTPS,25,2,"No"
								Print "No"
						End If	
            End If   
        Loop  
        If withRepeat Then  
            rep = Print ("This file does not exist:" & vbcr & vWaitFileORA & vbcr & vbcr & "Keep trying?", vbRetryCancel+vbExclamation, "File not found")  
            doAgain = (rep = vbRetry)  
        Else  
            doAgain = False  
        End If  
    Loop  
   fnWaitTillExistsORA = False  
End Function 
'***************
Function  fnWaitTillExistsSQL(vWaitInstallFilePathSQL)  
'''    ' Sleeps until the file exists  
'''    ' The polling interval will increase gradually, but never rises above MAX_WAITTIME  
'''    ' Times out after TIMEOUT msec. Will return false if caused by timeout.  
    Dim waittime, totalwaittime, rep, doAgain,withRepeat,fso,arr(2)
    Const INIT_WAITTIME = 20  
    Const MAX_WAITTIME = 1000  
    Const TIMEOUT = 50000  
    Const SLOPE = 1.1  
	withRepeat = True
''''    file = replace (file, Chr(34), "") 'remove double quotes from the input  
    doAgain  = True  
''''    Set WshShell = CreateObject( "WScript.Shell") 
	Set fso = createobject("Scripting.FileSystemObject")
    Do While doAgain  
        waittime = INIT_WAITTIME  
        totalwaittime = 0  
        Do While totalwaittime < TIMEOUT  
            waittime = Int (waittime * SLOPE)  
            If waittime>MAX_WAITTIME Then waittime=MAX_WAITTIME  
            totalwaittime = totalwaittime + waittime  
			Wait(waittime)
''''			**************************
'			vWaitInstallFilePathSQL = vInstallPath&vInstallSql
			vWaitInstallSQLLog = fnGetASingleFileInFolder(vWaitInstallFilePathSQL)
			vWaitFileSQL = vWaitInstallFilePathSQL&"\"&vWaitInstallSQLLog 
	
''			REM ******************* Verifying Migration Check  warning dialog box **********
            If fso.fileExists (vWaitFileSQL) Then
					fnWaitTillExistsSQL = True
				 
					Exit Function 
				else
						If  fnMigrationCheck = "Migration Check" Then
                            		fnWriteToExcel vTP_Path,vTPS,26,2,"Yes"
									Print "Yes"	
						Else	
									Print "No"						
'																		
									
						End If						
				End If   
        Loop  
        If withRepeat Then  
            rep = Print ("This file does not exist:" & vbcr & vWaitFileSQL & vbcr & vbcr & "Keep trying?", vbRetryCancel+vbExclamation, "File not found")  
            doAgain = (rep = vbRetry)  
        Else  
            doAgain = False  
        End If  
    Loop  
   fnWaitTillExistsSQL = False  
End Function 

'''fnSaveBLiExMngSetting()
Function fnSaveBLiExMngSetting()	
'''			SystemUtil.Run "C:\Program Files\Aligned Assets Limited\Symphony Symphony iExchange Manager\Symphony Symphony iExchange Settings.exe"
			print "fnSaveBLiExMngSetting working "
			SystemUtil.Run "C:\Program Files\Aligned Assets Limited\Symphony iExchange\Symphony Symphony iExchange Settings.exe"
						
			Window("Symphony Symphony iExchange").WinObject("Test Connection").Click
			
			Window("Symphony Symphony iExchange").Dialog("Symphony Symphony iExchange").WinButton("OK").Click		
			
			Window("Symphony Symphony iExchange").WinObject("Save").Click
End Function

Function fniExSetting(vSystem_Folder)
			Systemutil.Run "C:\Program Files\Aligned Assets Limited\iExchange\Symphony.iExchange.Settings.exe"
			Dialog("Symphony iExchange Settings").WinButton("OK").Click
			SwfWindow("Symphony iExchange Settings").SwfObject("utbcSettingsInfo").Click 92,15
			SwfWindow("Symphony iExchange Settings").SwfObject("uteSystemFolder").Click 103,8
			OptionalStep.SwfWindow("Symphony iExchange Settings").SwfEdit("uteSystemFolder_EmbeddableText").Set ""
			SwfWindow("Symphony iExchange Settings").SwfEdit("uteSystemFolder_EmbeddableText").Set vSystem_Folder
			SwfWindow("Symphony iExchange Settings").SwfObject("Save").Click 45,12
			SwfWindow("Symphony iExchange Settings").SwfObject("Close").Click 46,13
End Function

Function fnMigrationCheck()
					On error resume next	
					 fnMigrationCheck = Trim(Dialog("Migration Check").GetROProperty("regexpwndtitle")) 
'					 fnMigrationCheck = "Migration Check" 'for debugging purpose
					
End Function


Function fnGetTimeStampOfSpecificFile(vFileName)
			dim fso, file, lastUpdated
			Set fso = CreateObject("Scripting.FileSystemObject")
			set file = fso.GetFile (vFileName)
			lastUpdated =  file.DateLastModified
			fnGetTimeStampOfSpecificFile = lastUpdated
End Function



Function fnSearchPattern(vPattern, vString)
		Set a = new RegExp
        a.Pattern = vPattern
		a.IgnoreCase = TRUE
		a.Global = True
Set matches = a.Execute(vString)
For each match in Matches
	b= match.value
Next
'	print "matching value :"&b
	On error resume next
    fnSearchPattern =  b
    Print "Heloo smooth"
End Function
'*********************




'**********************

''Dim vTo:vTo = "pradeep.lalung@aligned-assets.co.uk"
''Dim vCC:vCC = "pradeep.lalung@aligned-assets.co.uk"
''Dim vTestEnvironment_ORA:vTestEnvironment_ORA = "ORA"
''fnSendReportOnMASTERBaseLineCreation vTo,vCC,vTestEnvironment_SQL,vExecutionTime,vSym_iX_Version1,vSym_iX_Service_Date,vSym_iX_Manager_Version1, _
''vSym_iX_Manager_Date,vSym_DB_Version1,vSym_DB_Build_Date,vMigrationCheck_SQL,vCsvFile_SQL,vImportReportFiles_SQL,vSP_Build_Date,vSP_Version1, _
''vSym_IXDb_Ver1,vSym_iX_DB_Date,vDBMigrationONOFF,vSinglePointTestONOFF,vIX_DBMigrationONOFF

