'Test Objectives : Copy and pasting required batch files and xls files from network drive
Option Explicit
Dim vStartTime: vStartTime = Time
Dim vMsgServiceStarted,vMsgServiceStopped,vServiceName,vSourceName,vComputerName,vStartedSuccessfully,vStoppedSuccessfully
Dim objWMIService,colLoggedEvents,colServiceList,objService,objEvent
Dim vOracle_Environment ,vSQL_Environment
Dim vServiceStarted,vServiceStopped
Const  vStarted = "Service Started"
Const  vStopped= "Service Stopped"

Dim vApp_Path,vTP_Path,vTP,vTPC,vTCD,vTCD1,vTCD2,vTCD3,vTCD4,vTSEC,vTPS,vAppStatus

Dim vInstallLog,vInstallLogPathORA,vInstallLogPathSQL,vInstallOra,vInstallSql,vInstallPath,vExtension,vLogName
Dim vWaitInstallORALog,vWaitInstallSQLLog,vWaitInstallFilePathORA,vWaitInstallFilePathSQL,vWaitFileSQL,vWaitFileORA 
Dim vMigrationCheck_ORA,vMigrationCheck_SQL
Dim vOraDBUpgrade,vSqlDBUpgrade,vExistControllerORA,vExistControllerSQL
Dim vSym_iX_Service,vSym_iX_Service1,vSym_iX_Service_Date,vSym_iX_Version,vSym_iX_Version1,vSym_iX_Manager,vSym_iX_Manager1,_
vSym_iX_Manager_Version,vSym_iX_Manager_Version1,vSym_iX_Manager_Date,vSym_DB_Build,vSym_DB_Build1,vSym_DB_Version,_
vSym_DB_Version1,vSym_DB_Build_Date,vAutoProcess_LogFile,vSym_iX_Ver_Paths,vCsvFile_ORA,vCsvFile_SQL

Dim vSym_IX_BldPath_Len,vSym_IX_Mng_BldPath_Len,vSym_BLD_BldPath_Len_ORA,vSym_BLD_BldPath_Len_SQL

Dim vTo,vCC,vSym_AllBlds_Folder,vAutoProcess_Log,vSym_IX_LogTxt_Path,vImportReportFiles_ORA,vImportReportFiles_SQL

Dim vCOUFiles,vCOUCombinedFolder,vCOUImportErrorFiles,vCOUImportErrorFilesSQL
Dim vLOFiles,vLOImportErrorFilesSQL,vLOImportErrorFiles

Dim vInitializationVBScript_SQL,vInitializationVBScript_ORA,vReplaceScriptName_SQL,vReplaceScriptName_ORA

Dim vSym_Path_ReplaceScriptName_SQL,vSP_Path_ReplaceScriptName_SQL,vSPReplaceScriptName_ORA

Dim vTestPlan_Completion
Dim k:k=2
Rem - Those below variable data are not being used as values are hard code in the code, but will be used in later phase
'Dim vSym_TP07_ORA,vSym_TP09_ORA,vSym_TP10_ORA,vSym_TP07_ReportORA,vSym_TP07_ReportORA,vSym_TP07_ReportORA
Dim vSym_TP07_SQL,vSym_TP09_SQL,vSym_TP10_SQL,vSym_TP11_SQL,vSym_TP13_SQL,vSym_TP16_SQL,vSym_TP17_SQL,vSym_TP19_SQL' vSym_TP09_ReportSQL,vSym_TP09_ReportSQL,vSym_TP10_ReportSQL

'***********************************************

Dim vTestEnvironment_ORA:vTestEnvironment_ORA = "ORACLE"
Dim vTestEnvironment_SQL:vTestEnvironment_SQL = "SQL"
'
'				****** STEP1 ***************
'				@ Open network drive and 
'				@ Copy test plans, CSV files, xml file and batch file to QTP/UFT Machine
				Systemutil.Run "C:\Automation\Sym_iEx\BatchFile\Generic\SQL\Updated_TestPlan_SQL.bat"
'				Wait(30)	
				

	
				
'				****** STEP2 ***************				
'''				@ Assigned  Environment variable  values dynamically from XML file *************
				Environment.LoadFromFile("C:\Automation\Sym_iEx\Xml_File\Sym_TP_Data_Generic.xml")
				vTPC = Environment.Value("vTPC")
				vTPS = Environment.Value("vTPS")
				Print "vTPS  :"&vTPS
				vTP_Path = Environment.Value("vTP_Path")
				print "vTP_Path  :"&vTP_Path
				print "vTPC :"&vTPC

					
		
		
		'				****** STEP3 ***************

		
					vMsgServiceStarted = Environment.Value("vMsgServiceStarted")
		
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
		'				
					vTPS = Environment.Value("vTPS")
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
					
'					vWaitInstallFilePathORA = Environment.Value("vWaitInstallFilePathORA")
					vWaitInstallFilePathSQL = Environment.Value("vWaitInstallFilePathSQL")
					
			
		
					vInstallLog = Environment.Value("vInstallLog")
					
					vInstallPath = Environment.Value("vInstallPath")
					vSym_IX_LogTxt_Path = Environment.Value("vSym_IX_LogTxt_Path")
					
					print "vSym_IX_LogTxt_Path  :"&vSym_IX_LogTxt_Path
					vAutoProcess_Log = Environment.Value("vAutoProcess_Log")
					
					print "vAutoProcess_Log  :"&vAutoProcess_Log
					
					vDefaultLenghtBldPath = Environment.Value("vDefaultLenghtBldPath")
					
					vCsvFile_ORA = Environment.Value("vCsvFile_ORA")
					vCsvFile_SQL = Environment.Value("vCsvFile_SQL")
'					*****
					vImportReportFiles_ORA = Environment.Value("vImportReportFiles_ORA")
					print "vImportReportFiles_ORA  :"&vImportReportFiles_ORA
					vImportReportFiles_SQL= Environment.Value("vImportReportFiles_SQL")
					print "vImportReportFiles_SQL  :"&vImportReportFiles_SQL
'					*****
					vSym_IX_BldPath_Len = Environment.Value("vSym_IX_BldPath_Len")
					vSym_IX_Mng_BldPath_Len = Environment.Value("vSym_IX_Mng_BldPath_Len")
					vSym_BLD_BldPath_Len_ORA = Environment.Value("vSym_BLD_BldPath_Len_ORA")
					vSym_BLD_BldPath_Len_SQL = Environment.Value("vSym_BLD_BldPath_Len_SQL")
					
					
'					******
					vCOUFiles = Environment.Value("vCOUFiles")
					print "vCOUFiles  :"&vCOUFiles
					vCOUCombinedFolder = Environment.Value("vCOUCombinedFolder")
					print "vCOUCombinedFolder  :"&vCOUCombinedFolder
					vCOUImportErrorFilesSQL = Environment.Value("vCOUImportErrorFilesSQL")
					print "vCOUImportErrorFilesSQL  :"&vCOUImportErrorFilesSQL
					vLOFiles = Environment.value("vLOFiles")
					print "vLOFiles  :"&vLOFiles 
					vLOImportErrorFiles = Environment.value("vLOImportErrorFiles") ' For Oracle
					vLOImportErrorFilesSQL = Environment.value("vLOImportErrorFilesSQL")
					print "vLOImportErrorFilesSQL  :"&vLOImportErrorFilesSQL 
					
					vInitializationVBScript_SQL = Environment.Value("vInitializationVBScript_SQL")
					print " vInitializationVBScript_SQL  :"&vInitializationVBScript_SQL
					vInitializationVBScript_ORA = Environment.Value("vInitializationVBScript_ORA")
					vReplaceScriptName_SQL = Environment.Value("vReplaceScriptName_SQL")
					print "vReplaceScriptName_SQL  :"&vReplaceScriptName_SQL
					vReplaceScriptName_ORA = Environment.Value("vReplaceScriptName_ORA")
					
					vSym_Path_ReplaceScriptName_SQL = Environment.Value("vSym_Path_ReplaceScriptName_SQL")
					print "vSym_Path_ReplaceScriptName_SQL  :"&vSym_Path_ReplaceScriptName_SQL
					
					vSP_Path_ReplaceScriptName_SQL = Environment.Value("vSP_Path_ReplaceScriptName_SQL")
					
					
					'fnDeleteFileInFolder vCOUFiles
					'fnDeleteFolder vCOUCombinedFolder
				    
		'				****** STEP4 ***************	
		'				@ Read the information from Test Plan
					vTo = fnReadFromExcel(vTP_Path,vTPS,12,2)
					vCC = fnReadFromExcel(vTP_Path,vTPS,13,2)
					
'					vSym_AllBlds_Folder = Trim(fnReadFromExcel(vTP_Path,vTPS,9,2))
		
					vSym_DB_Build = Trim(fnReadFromExcel(vTP_Path,vTPS,19,2)) ' "C:\Automation\Sym_iEx\AutomatedInstall\Symphony Bluelight Oracle 5.5.0.0.0.0.exe"			
			
					vOraDBUpgrade= fnReadFromExcel(vTP_Path,vTPS,15,2)
					print "vOraDBUpgrade  :"&vOraDBUpgrade
					vSqlDBUpgrade= fnReadFromExcel(vTP_Path,vTPS,16,2) ' 
					print "vSqlDBUpgrade  :"&vSqlDBUpgrade 
					vOracle_Environment  = fnReadFromExcel(vTP_Path,vTPS,23,2)
					print "vOracle_Environment  :"&vOracle_Environment
					vSQL_Environment = fnReadFromExcel(vTP_Path,vTPS,24,2)
					print "vSQL_Environment  :"&vSQL_Environment
					systemutil.Run "C:\Automation\Sym_iEx\BatchFile\Generic\Sym_iEx_AccSet.bat"

'When the script will find YES in the TestPlan >> TestPlan_Control (sheet), it will move to next testplan in the queue
Do while fnReadFromExcel(vTP_Path,vTPC,k,6) <> "END"  and fnReadFromExcel(vTP_Path,vTPC,k,6) = "Yes"
				'*******************New Instructural change
				print "Start K  :"&k
				k=k+1
				If fnReadFromExcel(vTP_Path,vTPC,k,6) = "No" Then
					vTestPlan = fnReadFromExcel(vTP_Path,vTPC,k,6)
					print "Test Plan in Progress :"&vTestPlan
					print "Value of K1234 :"&k
					
				
											Select Case k
												Case 3	
'												*******************
											       Rem @ Copy all the necessary Batch files from Network drive to QTP machine*****
'											       
											       print "Staring - Sym_TP07"
											       	vSym_TP07_SQL = Environment.Value("vSym_TP07_SQL")
'											       	vSym_TP07_ReportSQL = Environment.Value("vSym_TP07_ReportSQL")
'											       	print "vSym_TP07_SQL  :"&vSym_TP07_SQL
'											       	REM - fnPreTest_TP07_SQL vSQL_Environment,vSym_TP07_SQL
											       	'Bring all generic updated batch files to QTP machine - Currently commentting out as need some validation
'											       	Systemutil.Run "C:\Automation\Sym_iEx\BatchFile\Generic\SQL\AllBatch4Sym_TP07_SQL.bat"
																										
													
																
														REM - fnSym_TP07_T002_SQL() ' No Bulk import , Not implemented as current csv is very old
														fnSym_TP07_T003_SQL() ' Bulk import
														
														fnReport_Sym_TP07_SQL
														print "Ending - Sym_TP07"
'														******************
												Case 4
														Rem @ Copy all the necessary Batch files from Network drive to QTP machine*****
														vSym_TP09_SQL = Environment.Value("vSym_TP09_SQL")
														print "vSym_TP09_SQL :"&vSym_TP09_SQL
'														REM - fnPreTest_TP09_SQL vSQL_Environment,vSym_TP09_SQL
'														Systemutil.Run "C:\Automation\Sym_iEx\BatchFile\Generic\SQL\AllBatch4Sym_TP09_SQL.bat"
														
														print "Sym_TP09"	
'														fnSym_TP09_T006_BN_SQL()' Skipping for time being
														fnSym_TP09_T006_SQL()
'														fnSym_TP09_T007_BN_SQL() ' Skipping for time being
														fnSym_TP09_T007_SQL()
																										
														fnReport_Sym_TP09_SQL
												Case 5
														Rem @ Copy all the necessary Batch files from Network drive to QTP machine*****
														vSym_TP10_SQL = Environment.Value("vSym_TP10_SQL")
														REM - fnPreTest_TP10_SQL vSQL_Environment,vSym_TP10_SQL
'														Systemutil.Run "C:\Automation\Sym_iEx\BatchFile\Generic\SQL\AllBatch4Sym_TP10_SQL.bat"
														
														print "Sym_TP10"
														fnSym_TP10_PreTest_SQL()
											
														fnReport_Sym_TP10_PreTest_SQL														
												Case 6	
														
														print "Sym_TP11"		
														Rem @ Copy all the necessary Batch files from Network drive to QTP machine*****
														vSym_TP11_SQL = Environment.Value("vSym_TP11_SQL")
														REM - fnPreTest_TP11_SQL vSQL_Environment,vSym_TP11_SQL
'														Systemutil.Run "C:\Automation\Sym_iEx\BatchFile\Generic\SQL\AllBatch4Sym_TP11_SQL.bat"
														print "Sym_TP11"
														fnSym_TP11_T001_SQL
														fnSym_TP11_T002_SQL
														fnSym_TP11_T003_SQL	
														
														fnReport_Sym_TP11_SQL																											
																											
												Case else
														Print "Run out of numbers"
											End Select
'									**************************
				End If

	If fnReadFromExcel(vTP_Path,vTPC,k,6) = "No" Then
		print "Hello - Test Plan is done"
		fnWriteToExcel vTP_Path,vTPC,k,6, "Yes"
		Rem ### Copy the updated testplan in R:Drive from where next test in schedule will pick up
		Systemutil.Run "C:\Automation\Sym_iEx\BatchFile\Generic\SQL\IX_TestPlan_Update2Yes_SQL.bat"
		wait(5)
				' The following batch will copy and paste the test plan in 
		' "R:\Symphony Bluelight iExchange\Testing\Automation\AddressBase\Run_Completion_SQL "as an input for powershell script to do next task
		' It will also update the orginal Test plan i
		' R:\Symphony Bluelight iExchange\Testing\Automation\AddressBase\TestPlan_SQL so that next Test Plan will run based on
		' YES control present in the Sym_TP01_NSG_Core>TestPlan_Control sheet
		 Rem ### After finishing T002,T003 & T004 - Output is stored in R:Drive for powershell to kick to next testplan(BaselineMaster)
		SystemUtil.Run "C:\Automation\Sym_iEx\BatchFile\Generic\SQL\Successful_Output_SQL.bat"
		Exit do
	End If
Loop
'*************** 2. Hand over the control for SINGLEPOINT TEST PLAN*****
'*******************Update QTPLauncher file to Switch and hand over the control for Bluelight iXchange Test Plan after creation of Baseline Master ****
'Print "vInitializationVBScript_SQL  :"&vInitializationVBScript_SQL
'print "vSym_Path_ReplaceScriptName_SQL :"&vSym_Path_ReplaceScriptName_SQL
'Print "vSP_Path_ReplaceScriptName_SQL  :"&vSP_Path_ReplaceScriptName_SQL
fnReplaceLineWithSpecificValue vInitializationVBScript_SQL,vSym_Path_ReplaceScriptName_SQL,vSP_Path_ReplaceScriptName_SQL,vInitializationVBScript_SQL

'*************** Start PreTest****
Function fnPreTest_TP07_SQL(vSQL_Environment,vPreTest_Bat_Recipient)
	
Print "Enter - PreTest after Environment is YES"
		If vSQL_Environment  = "Yes" then
					    Print "Starting PreTest after Environment is YES"
						Rem - Clear the Event Viewer list ******
						Systemutil.Run "C:\Automation\Sym_iEx\BatchFile\ClearEventViewerListApplication.bat"
						REM - Saving Bluelight iExchange manager setting*****
						fnSaveBLiExMngSetting()
		'				****** STEP8 ***************
		'				@ First Phase of Service start ,stop and clear Event Viewer BEFORE import				
						
		'					************** 8a. Start the service *********************
						wait(2)
						Systemutil.Run "C:\Automation\Sym_iEx\BatchFile\Sym_IX_StartService.bat"
						wait(5)
		'					*************** 8b. ***The following function only comes out from executing when find TRUE ***
						vStartedSuccessfully =  fnWaitTillExistsStarted(vMsgServiceStarted,vSourceName)
						'*********************************
						If vStartedSuccessfully = True then
											REM *************** 8c. Stop the service *********
											wait(2)
											Systemutil.Run "C:\Automation\Sym_iEx\BatchFile\Sym_IX_StopService.bat"
											wait(5)
											REM *************** 8d. Make the QTP to wait for the service to Stop and provide TRUE once completes *********
											vStoppedSuccessfully =  fnWaitTillExistsStopped(vMsgServiceStopped,vSourceName)
											REM *************** 8e. 
											If  vStoppedSuccessfully = True  Then   
		
		
						'					Rem **************** Save the event log file for that application **************
'											SystemUtil.Run "C:\Automation\Sym_iEx\BatchFile\SaveEventLogApplication_SQL.bat"
						'					*********  Save the PreTest Recipient files in R:Drive ***************
'											SystemUtil.Run "C:\Automation\Sym_iEx\BatchFile\Sym_TP07_COU_Recipient_ORA_PreTest.bat"
											SystemUtil.Run "C:\Automation\Sym_iEx\BatchFile\IX_TP07\SQL\IX_TP07_COU_Recipient_SQL_PreTest.bat"
											REM **************** 3. Clear the Event Viewer list ******
											Systemutil.Run "C:\Automation\Sym_iEx\BatchFile\ClearEventViewerListSystem.bat"
											Print "1. Clear the event view list ****"
											Print "2. Stop the service"
											End If
						End if
		End If
End Function
Function fnPreTest_TP09_SQL(vSQL_Environment,vPreTest_Bat_Recipient)
	
Print "Enter - PreTest after Environment is YES - fnPreTest_TP09_SQL"
		If vSQL_Environment  = "Yes" then
					    Print "Starting PreTest after Environment is YES"
						Rem - Clear the Event Viewer list ******
						Systemutil.Run "C:\Automation\Sym_iEx\BatchFile\ClearEventViewerListApplication.bat"
						REM - Saving Bluelight iExchange manager setting*****
						fnSaveBLiExMngSetting()
		'				****** STEP8 ***************
		'				@ First Phase of Service start ,stop and clear Event Viewer BEFORE import				
						
		'					************** 8a. Start the service *********************
						wait(2)
						Systemutil.Run "C:\Automation\Sym_iEx\BatchFile\Sym_IX_StartService.bat"
						Wait(5)
		'					*************** 8b. ***The following function only comes out from executing when find TRUE ***
						vStartedSuccessfully =  fnWaitTillExistsStarted(vMsgServiceStarted,vSourceName)
						'*********************************
						If vStartedSuccessfully = True then
											REM *************** 8c. Stop the service *********
											wait(2)
											Systemutil.Run "C:\Automation\Sym_iEx\BatchFile\Sym_IX_StopService.bat"
											Wait(5)
											REM *************** 8d. Make the QTP to wait for the service to Stop and provide TRUE once completes *********
											vStoppedSuccessfully =  fnWaitTillExistsStopped(vMsgServiceStopped,vSourceName)
											REM *************** 8e. 
											If  vStoppedSuccessfully = True  Then   
		
		
						'					Rem **************** Save the event log file for that application **************
'											SystemUtil.Run "C:\Automation\Sym_iEx\BatchFile\SaveEventLogApplication_SQL.bat"
						'					*********  Save the PreTest Recipient files in R:Drive ***************
'											SystemUtil.Run "C:\Automation\Sym_iEx\BatchFile\IX_TP07_COU_Recipient_ORA_PreTest.bat"
											SystemUtil.Run "C:\Automation\Sym_iEx\BatchFile\IX_TP09\SQL\IX_TP09_LO_Recipient_SQL_PreTest.bat"
											REM **************** 3. Clear the Event Viewer list ******
											Systemutil.Run "C:\Automation\Sym_iEx\BatchFile\ClearEventViewerListSystem.bat"
											Print "1. Clear the event view list ****"
											Print "2. Stop the service"
											End If
						End if
		End If
End Function
Function fnPreTest_TP10_SQL(vSQL_Environment,vPreTest_Bat_Recipient)
	
Print "Enter - PreTest after Environment is YES - fnPreTest_TP10_SQL"
		If vSQL_Environment  = "Yes" then
					    Print "Starting PreTest after Environment is YES"
						Rem - Clear the Event Viewer list ******
						Systemutil.Run "C:\Automation\Sym_iEx\BatchFile\ClearEventViewerListApplication.bat"
						REM - Saving Bluelight iExchange manager setting*****
						fnSaveBLiExMngSetting()
		'				****** STEP8 ***************
		'				@ First Phase of Service start ,stop and clear Event Viewer BEFORE import				
						
		'					************** 8a. Start the service *********************
						Wait(2)
						Systemutil.Run "C:\Automation\Sym_iEx\BatchFile\Sym_IX_StartService.bat"
						Wait(5)
		'					*************** 8b. ***The following function only comes out from executing when find TRUE ***
						vStartedSuccessfully =  fnWaitTillExistsStarted(vMsgServiceStarted,vSourceName)
						'*********************************
						If vStartedSuccessfully = True then
											REM *************** 8c. Stop the service *********
											Wait(2)
											Systemutil.Run "C:\Automation\Sym_iEx\BatchFile\Sym_IX_StopService.bat"
											Wait(5)
											REM *************** 8d. Make the QTP to wait for the service to Stop and provide TRUE once completes *********
											vStoppedSuccessfully =  fnWaitTillExistsStopped(vMsgServiceStopped,vSourceName)
											REM *************** 8e. 
											If  vStoppedSuccessfully = True  Then   
		
		
						'					Rem **************** Save the event log file for that application **************
'											SystemUtil.Run "C:\Automation\Sym_iEx\BatchFile\SaveEventLogApplication_SQL.bat"
						'					*********  Save the PreTest Recipient files in R:Drive ***************
'											SystemUtil.Run "C:\Automation\Sym_iEx\BatchFile\IX_TP07_COU_Recipient_ORA_PreTest.bat"
											SystemUtil.Run "C:\Automation\Sym_iEx\BatchFile\IX_TP10\SQL\IX_TP10_LO_Recipient_SQL_PreTest.bat"
											REM **************** 3. Clear the Event Viewer list ******
											Systemutil.Run "C:\Automation\Sym_iEx\BatchFile\ClearEventViewerListSystem.bat"
											Print "1. Clear the event view list ****"
											Print "2. Stop the service"
											End If
						End if
		End If
End Function
'************
Function fnPreTest_TP11_SQL(vSQL_Environment,vPreTest_Bat_Recipient)
	
Print "Enter - PreTest after Environment is YES - fnPreTest_TP11_SQL"
		If vSQL_Environment  = "Yes" then
					    Print "Starting PreTest after Environment is YES"
						Rem - Clear the Event Viewer list ******
						Systemutil.Run "C:\Automation\Sym_iEx\BatchFile\ClearEventViewerListApplication.bat"
						REM - Saving Bluelight iExchange manager setting*****
						fnSaveBLiExMngSetting()
		'				****** STEP8 ***************
		'				@ First Phase of Service start ,stop and clear Event Viewer BEFORE import				
						
		'					************** 8a. Start the service *********************
						Wait(2)
						Systemutil.Run "C:\Automation\Sym_iEx\BatchFile\Sym_IX_StartService.bat"
						Wait(5)
		'					*************** 8b. ***The following function only comes out from executing when find TRUE ***
						vStartedSuccessfully =  fnWaitTillExistsStarted(vMsgServiceStarted,vSourceName)
						'*********************************
						If vStartedSuccessfully = True then
											REM *************** 8c. Stop the service *********
											wait(2)
											Systemutil.Run "C:\Automation\Sym_iEx\BatchFile\Sym_IX_SStopService.bat"
											wait(5)
											REM *************** 8d. Make the QTP to wait for the service to Stop and provide TRUE once completes *********
											vStoppedSuccessfully =  fnWaitTillExistsStopped(vMsgServiceStopped,vSourceName)
											REM *************** 8e. 
											If  vStoppedSuccessfully = True  Then   
		
		
						'					Rem **************** Save the event log file for that application **************
'											SystemUtil.Run "C:\Automation\Sym_iEx\BatchFile\SaveEventLogApplication_SQL.bat"
						'					*********  Save the PreTest Recipient files in R:Drive ***************
'											SystemUtil.Run "C:\Automation\Sym_iEx\BatchFile\IX_TP07_COU_Recipient_ORA_PreTest.bat"
											SystemUtil.Run "C:\Automation\Sym_iEx\BatchFile\IX_TP11\SQL\IX_TP11_LO_Recipient_SQL_PreTest.bat"
											REM **************** 3. Clear the Event Viewer list ******
											Systemutil.Run "C:\Automation\Sym_iEx\BatchFile\ClearEventViewerListSystem.bat"
											Print "1. Clear the event view list ****"
											Print "2. Stop the service"
											End If
						End if
		End If
End Function

'************End PreTest****



'********************************************************************************End of Sym_TP07 ******************

		'*************	STEP 10 Reporting *************		
Function fnReport_Sym_TP07_SQL()	
		
		Dim vEndTime:vEndTime = Time
		Dim vTimeTaken,vExecutionTime,vMigrationCheck
		vTimeTaken = vEndTime - vStartTime
		vExecutionTime = fnExecutionTime(vTimeTaken)
		
		vEndTime = Time
		vTimeTaken = vEndTime - vStartTime
		vExecutionTime = fnExecutionTime(vTimeTaken)
		'************ Send the above information via email to respective recipients/Stakeholders in the Test Plan
		 fnSend_TestResult_Sym_TP07_SQL vTo,vCC,vTestEnvironment_SQL,vExecutionTime,vSym_iX_Version1,vSym_iX_Service_Date,vSym_iX_Manager_Version1,_
		vSym_iX_Manager_Date,vSym_DB_Version1,vSym_DB_Build_Date,vMigrationCheck_SQL,vCsvFile_SQL,vImportReportFiles_SQL
		
End Function
Function fnReport_Sym_TP09_SQL()	
		
		Dim vEndTime:vEndTime = Time
		Dim vTimeTaken,vExecutionTime,vMigrationCheck
		vTimeTaken = vEndTime - vStartTime
		vExecutionTime = fnExecutionTime(vTimeTaken)
		
		vEndTime = Time
		vTimeTaken = vEndTime - vStartTime
		vExecutionTime = fnExecutionTime(vTimeTaken)
		'************ Send the above information via email to respective recipients/Stakeholders in the Test Plan
		 fnSend_TestResult_Sym_TP09_SQL vTo,vCC,vTestEnvironment_SQL,vExecutionTime,vSym_iX_Version1,vSym_iX_Service_Date,vSym_iX_Manager_Version1,_
		 vSym_iX_Manager_Date,vSym_DB_Version1,vSym_DB_Build_Date,vMigrationCheck_SQL,vCsvFile_SQL,vImportReportFiles_SQL
		
End Function
Function fnReport_Sym_TP10_PreTest_SQL()	
		
		Dim vEndTime:vEndTime = Time
		Dim vTimeTaken,vExecutionTime,vMigrationCheck
		vTimeTaken = vEndTime - vStartTime
		vExecutionTime = fnExecutionTime(vTimeTaken)
		
		vEndTime = Time
		vTimeTaken = vEndTime - vStartTime
		vExecutionTime = fnExecutionTime(vTimeTaken)
		'************ Send the above information via email to respective recipients/Stakeholders in the Test Plan
		fnSend_TestResult_Sym_TP10_SQL vTo,vCC,vTestEnvironment_SQL,vExecutionTime,vSym_iX_Version1,vSym_iX_Service_Date,vSym_iX_Manager_Version1,_
		 vSym_iX_Manager_Date,vSym_DB_Version1,vSym_DB_Build_Date,vMigrationCheck_SQL,vCsvFile_SQL,vImportReportFiles_SQL
		 
		
End Function
'**************
Function fnReport_Sym_TP11_SQL()	
		
		Dim vEndTime:vEndTime = Time
		Dim vTimeTaken,vExecutionTime,vMigrationCheck
		vTimeTaken = vEndTime - vStartTime
		vExecutionTime = fnExecutionTime(vTimeTaken)
		
		vEndTime = Time
		vTimeTaken = vEndTime - vStartTime
		vExecutionTime = fnExecutionTime(vTimeTaken)
		'************ Send the above information via email to respective recipients/Stakeholders in the Test Plan
		 fnSend_TestResult_Sym_TP11_SQL vTo,vCC,vTestEnvironment_SQL,vExecutionTime,vSym_iX_Version1,vSym_iX_Service_Date,vSym_iX_Manager_Version1,_
		 vSym_iX_Manager_Date,vSym_DB_Version1,vSym_DB_Build_Date,vMigrationCheck_SQL,vCsvFile_SQL,vImportReportFiles_SQL	 		 
		 
	
		
End Function
'*************



''************************ End of the Script *******************
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

Function fnImportBuild_Sym_TP07()			
		RunAction "Action1 [T001_LLPG_WithoutBulkImport]", oneIteration
End Function

Function fnImportBuild_Sym_Bulk_TP07()

		RunAction "Action1 [T001_LLPG_WithBulkImport]", oneIteration			
End Function

Function fnImportBuild_Sym_TP09()
		RunAction "Action1 [T001_LLPG_WithoutBulkImport]", oneIteration
End Function

Function fnImportBuild_Sym_TP10()
			RunAction "Action1 [T001_LLPG_WithoutBulkImport_FlagRecords]", oneIteration
End Function

Function fnImportBuild_Sym_TP11()
		RunAction "Action1 [T001_LLPG_WithoutBulkImport]", oneIteration
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
'''			SystemUtil.Run "C:\Program Files\Aligned Assets Limited\Symphony Bluelight iExchange Manager\Symphony Bluelight iExchange Settings.exe"
			print "fnSaveBLiExMngSetting working "
			SystemUtil.Run "C:\Program Files\Aligned Assets Limited\Bluelight iExchange\Symphony Bluelight iExchange Settings.exe"
						
			Window("Symphony Bluelight iExchange").WinObject("Test Connection").Click
			
			Window("Symphony Bluelight iExchange").Dialog("Symphony Bluelight iExchange").WinButton("OK").Click		
			
			Window("Symphony Bluelight iExchange").WinObject("Save").Click
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



'***********
		'				@ Import of CSV file start - Sym_TP09 ******************										
		'				****** STEP9 ***************
		'				@ Import of CSV file start - Sym_TP07 ******************										

Function fnSym_TP07_T002_SQL() ' No Bulk Import
											print "Test Case TP07_T002 Starting"
											REM - # Copy csv file from generic common location(R:Drive) to specific location (R:Drive)
											Systemutil.Run "C:\Automation\Sym_iEx\BatchFile\IX_TP07\Generic\IX_TP07_T002_COU_CSV.bat"
											wait(5)
											REM - # Copy csv file from specific location(R:Drive) to QTP machine
											Systemutil.Run "C:\Automation\Sym_iEx\BatchFile\IX_TP07\Generic\IX_TP07_T00234_COU_CSV_Local.bat"
											wait(5)
											Print "Import process started"
											fnImportBuild_Sym_TP07
											Print "Import process successfully completed"
											REM - Copy all the import report into network Drive 
											Print "Copy all the import report into network Drive_CopyImportReportSym_TP07_T002_SQL"
											fnCopyImportReportSym_TP07_T002_SQL()
											wait(5)
											Print "Delete all the import report,COU csv files and combined folder"
'											fnDeleteFileInFolder vCOUFiles
'											fnDeleteFolder vCOUCombinedFolder
'											fnDeleteFileInFolder vCOUImportErrorFiles
											fnDeleteFileFolderSQL
											
											
											REM **************************** Second Phase of Service start ,stop and clear  Event Viewer AFTER import ********
											REM ************** 1. Start the service *********************
																Rem - Clear the Event Viewer list ******
											print "SECOND PHASE - CLEAN Event VIEWER SQL"
											Systemutil.Run "C:\Automation\Sym_iEx\BatchFile\ClearEventViewerListApplication.bat"
		'										*****************
											REM - Saving Bluelight iExchange manager setting*****
'											fnSaveBLiExMngSetting()
											
		'										**************
											'fnStartService vServiceName,vComputerName
											wait(2)
											Systemutil.Run "C:\Automation\Sym_iEx\BatchFile\Sym_IX_StartService.bat"
											wait(5)
											'********************************
											vStartedSuccessfully =  fnWaitTillExistsStarted(vMsgServiceStarted,vSourceName)
		'										************
														If vStartedSuccessfully = True then
															REM *************** 2. Stop the service *********
										'					fnStopService vServiceName,vComputerName
															wait(2)
															Systemutil.Run "C:\Automation\Sym_iEx\BatchFile\Sym_IX_StopService.bat"
															Wait(5)
															
															vStoppedSuccessfully =  fnWaitTillExistsStopped(vMsgServiceStopped,vSourceName)
															If vStoppedSuccessfully = True Then
																	Print "Copying the recipient folder to network drive SQL"
																	Rem ### Copying recipient files in R:Drive
																	SystemUtil.Run "C:\Automation\Sym_iEx\BatchFile\IX_TP07\SQL\IX_TP07_T002_COU_Recipient_SQL.bat"
''																	
															End If
														End If
															
		'										***********
		print "Test Case TP07_T002 END"
End Function
Function fnSym_TP07_T003_SQL() ' Bulk import
											print "Test Case TP07_T003 Starting"
											Systemutil.Run "C:\Automation\Sym_iEx\BatchFile\IX_TP07\Generic\IX_TP07_T003_COU_CSV.bat"
											wait(5)
											Systemutil.Run "C:\Automation\Sym_iEx\BatchFile\IX_TP07\Generic\IX_TP07_T0023_COU_CSV_Local.bat"
											wait(5)
											Print "Import process started"
											fnImportBuild_Sym_Bulk_TP07
											Print "Import process successfully completed"
											REM - Copy all the import report into network Drive 
											Print "Copy all the import report into network Drive_CopyImportReportSym_TP07_T003_SQL"
											fnCopyImportReportSym_TP07_T003_SQL()											
											wait(5)
											Print "Delete all the import report,COU csv files and combined folder"

											fnDeleteFileFolderSQL
											REM **************************** Second Phase of Service start ,stop and clear  Event Viewer AFTER import ********
											REM ************** 1. Start the service *********************
																Rem - Clear the Event Viewer list ******
											print "SECOND PHASE - CLEAN Event VIEWER SQL"
											Systemutil.Run "C:\Automation\Sym_iEx\BatchFile\ClearEventViewerListApplication.bat"
		'										*****************
											REM - Saving Bluelight iExchange manager setting*****
'											fnSaveBLiExMngSetting()
											
'												**************

											wait(2)
											Systemutil.Run "C:\Automation\Sym_iEx\BatchFile\Sym_IX_StartService.bat"
											Wait(5)
											'********************************
											vStartedSuccessfully =  fnWaitTillExistsStarted(vMsgServiceStarted,vSourceName)
		'										************
														If vStartedSuccessfully = True then
															REM *************** 2. Stop the service *********
										'					fnStopService vServiceName,vComputerName
															wait(2)
															Systemutil.Run "C:\Automation\Sym_iEx\BatchFile\Sym_IX_StopService.bat"
															wait(5)
															vStoppedSuccessfully =  fnWaitTillExistsStopped(vMsgServiceStopped,vSourceName)
															If vStoppedSuccessfully = True Then
																	Print "Copying the recipient folder to network drive SQL"
																	Rem ### Copying recipient files in R:Drive
																	SystemUtil.Run "C:\Automation\Sym_iEx\BatchFile\IX_TP07\SQL\IX_TP07_T003_COU_Recipient_SQL.bat"
'		
															End If
														End If
															
		'										***********
		print "Test Case TP07_T003 END"

End Function
Function fnSym_TP09_T006_SQL() ' No Bulk update
											print "Test Case TP09_T002 Starting"
											Print "Delete all the import report,LO csv files and combined folder from previous ImportT001"
											fnDeleteFileFolderSQL
											Systemutil.Run "C:\Automation\Sym_iEx\BatchFile\IX_TP09\Generic\IX_TP09_T006_COU_CSV.bat"
											wait(5)
											Systemutil.Run "C:\Automation\Sym_iEx\BatchFile\IX_TP09\Generic\IX_TP09_T0067_COU_CSV_Local.bat"
											wait(5)
											Print "Import process started"
											fnImportBuild_Sym_TP09
											Print "Import process successfully completed"
											REM - Copy all the import report into network Drive 
											Print "Copy all the import report into network Drive_CopyImportReportSym_TP09_T002_SQL"
											fnCopyImportReportSym_TP09_T006_SQL()
											wait(5)
'											fnDeleteFileFolderSQL
										
											
											REM **************************** Second Phase of Service start ,stop and clear  Event Viewer AFTER import ********
											REM ************** 1. Start the service *********************
																Rem - Clear the Event Viewer list ******
											print "SECOND PHASE - CLEAN Event VIEWER SQL"
'											Systemutil.Run "C:\Automation\Sym_iEx\BatchFile\ClearEventViewerListApplication.bat"
'		'										*****************
'											REM - Saving Bluelight iExchange manager setting*****
'											fnSaveBLiExMngSetting()
'											
'		'										**************
											'fnStartService vServiceName,vComputerName
											wait(2)
											Systemutil.Run "C:\Automation\Sym_iEx\BatchFile\Sym_IX_StartService.bat"
											wait(5)
											'********************************
											vStartedSuccessfully =  fnWaitTillExistsStarted(vMsgServiceStarted,vSourceName)
		'										************
														If vStartedSuccessfully = True then
															REM *************** 2. Stop the service *********
										'					fnStopService vServiceName,vComputerName
															wait(2)
															Systemutil.Run "C:\Automation\Sym_iEx\BatchFile\Sym_IX_StopService.bat"
															wait(5)
															vStoppedSuccessfully =  fnWaitTillExistsStopped(vMsgServiceStopped,vSourceName)
															If vStoppedSuccessfully = True Then
																	Print "Copying the recipient folder to network drive SQL"
																	Rem ### Copying recipient files in R:Drive
																	SystemUtil.Run "C:\Automation\Sym_iEx\BatchFile\IX_TP09\SQL\IX_TP09_T006_COU_Recipient_SQL.bat"
''						
															End If
														End If
															
		'										***********
		print "Test Case TP09_T002 END"
End Function

Function fnSym_TP09_T007_SQL() ' No Bulk update
											print "Test Case TP09_T007 Starting"
											Print "Delete all the import report,LO csv files and combined folder from previous ImportT001"
											fnDeleteFileFolderSQL
											Systemutil.Run "C:\Automation\Sym_iEx\BatchFile\IX_TP09\Generic\IX_TP09_T007_COU_CSV.bat"
											wait(5)
											Systemutil.Run "C:\Automation\Sym_iEx\BatchFile\IX_TP09\Generic\IX_TP09_T0067_COU_CSV_Local.bat"
											wait(5)
											Print "Import process started"
											fnImportBuild_Sym_TP09
											Print "Import process successfully completed"
											REM - Copy all the import report into network Drive 
											Print "Copy all the import report into network Drive_CopyImportReportSym_TP09_T002_SQL"
											fnCopyImportReportSym_TP09_T007_SQL()
											wait(5)
											fnDeleteFileFolderSQL
										
											
											REM **************************** Second Phase of Service start ,stop and clear  Event Viewer AFTER import ********
											REM ************** 1. Start the service *********************
																Rem - Clear the Event Viewer list ******
											print "SECOND PHASE - CLEAN Event VIEWER SQL"
'											Systemutil.Run "C:\Automation\Sym_iEx\BatchFile\ClearEventViewerListApplication.bat"
'		'										*****************
'											REM - Saving Bluelight iExchange manager setting*****
'											fnSaveBLiExMngSetting()
'											
'		'										**************
											'fnStartService vServiceName,vComputerName
											wait(2)
											Systemutil.Run "C:\Automation\Sym_iEx\BatchFile\Sym_IX_StartService.bat"
											wait(5)
											'********************************
											vStartedSuccessfully =  fnWaitTillExistsStarted(vMsgServiceStarted,vSourceName)
		'										************
														If vStartedSuccessfully = True then
															REM *************** 2. Stop the service *********
										'					fnStopService vServiceName,vComputerName
															wait(2)
															Systemutil.Run "C:\Automation\Sym_iEx\BatchFile\Sym_IX_StopService.bat"
															wait(5)
															vStoppedSuccessfully =  fnWaitTillExistsStopped(vMsgServiceStopped,vSourceName)
															If vStoppedSuccessfully = True Then
																	Print "Copying the recipient folder to network drive SQL"
																	Rem ### Copying recipient files in R:Drive
																	SystemUtil.Run "C:\Automation\Sym_iEx\BatchFile\IX_TP09\SQL\IX_TP09_T007_COU_Recipient_SQL.bat"
''						
															End If
														End If
															
		'										***********
		print "Test Case TP09_T007 END"
End Function

Function fnSym_TP10_PreTest_SQL() ' No Bulk update
											print "Sym_TP10_PreTest_SQL Starting"
											Print "Delete all the import report,LO csv files and combined folder from previous ImportT001"
											fnDeleteFileFolderSQL
											Systemutil.Run "C:\Automation\Sym_iEx\BatchFile\IX_TP10\Generic\IX_TP10_PreTest_COU_CSV.bat"
											wait(5)
											Systemutil.Run "C:\Automation\Sym_iEx\BatchFile\IX_TP10\Generic\IX_TP10_PreTest_COU_CSV_Local.bat"
											wait(5)
											Print "Import process started"
											fnImportBuild_Sym_TP10
											Print "Import process successfully completed"
											REM - Copy all the import report into network Drive 
											Print "Copy all the import report into network Drive_CopyImportReportSym_TP09_T002_SQL"
											fnCopyImportReportSym_TP10_PreTest_SQL()
											wait(5)
											fnDeleteFileFolderSQL
										
'****************************** NO service Operation ****************************************************************************************														
											REM **************************** Second Phase of Service start ,stop and clear  Event Viewer AFTER import ********
											REM ************** 1. Start the service *********************
																Rem - Clear the Event Viewer list ******
'''											print "SECOND PHASE - CLEAN Event VIEWER SQL"
''''											Systemutil.Run "C:\Automation\Sym_iEx\BatchFile\ClearEventViewerListApplication.bat"
''''		'										*****************
''''											REM - Saving Bluelight iExchange manager setting*****
''''											fnSaveBLiExMngSetting()
''''											
''''		'										**************
'''											'fnStartService vServiceName,vComputerName
'''											wait(2)
'''											Systemutil.Run "C:\Automation\Sym_iEx\BatchFile\Sym_IX_StartService.bat"
'''											wait(5)
'''											'********************************
'''											vStartedSuccessfully =  fnWaitTillExistsStarted(vMsgServiceStarted,vSourceName)
'''		'										************
'''														If vStartedSuccessfully = True then
'''															REM *************** 2. Stop the service *********
'''										'					fnStopService vServiceName,vComputerName
'''															wait(2)
'''															Systemutil.Run "C:\Automation\Sym_iEx\BatchFile\Sym_IX_StopService.bat"
'''															wait(5)
'''															vStoppedSuccessfully =  fnWaitTillExistsStopped(vMsgServiceStopped,vSourceName)
'''															If vStoppedSuccessfully = True Then
'''																	Print "Copying the recipient folder to network drive SQL"
'''																	Rem ### Copying recipient files in R:Drive
'''																	SystemUtil.Run "C:\Automation\Sym_iEx\BatchFile\IX_TP09\SQL\IX_TP09_T007_COU_Recipient_SQL.bat"
'''''						
'''															End If
'''														End If
															
		'										***********
		print "Test Case TP10_Pretest END"
End Function

'******************* Sym_TP11**********

Function fnSym_TP11_T001_SQL()
								print "Test Case TP11_T001_Starting"
								Systemutil.Run "C:\Automation\Sym_iEx\BatchFile\IX_TP11\Generic\IX_TP11_T001_COU_CSV.bat"
								wait(5)
								Systemutil.Run "C:\Automation\Sym_iEx\BatchFile\IX_TP11\Generic\IX_TP11_T0013_COU_CSV_Local.bat"
								wait(5)
								Print "Import process started"
								fnImportBuild_Sym_TP11
								Print "Import process successfully completed"
								REM - Copy all the import report into network Drive 
								Print "Copy all the import report into network Drive_CopyImportReportSym_TP11_T001_SQL"
								fnCopyImportReportSym_TP11_T001_SQL()
								wait(5)
								Print "Delete all the import report,LO csv files and combined folder"
								fnDeleteFileFolderSQL
								
								
								REM **************************** Second Phase of Service start ,stop and clear  Event Viewer AFTER import ********
								REM ************** 1. Start the service *********************
													Rem - Clear the Event Viewer list ******
								print "SECOND PHASE - CLEAN Event VIEWER SQL"
								Systemutil.Run "C:\Automation\Sym_iEx\BatchFile\ClearEventViewerListApplication.bat"
'										*****************
								REM - Saving Bluelight iExchange manager setting*****
'								fnSaveBLiExMngSetting()
'								
''										**************
								'fnStartService vServiceName,vComputerName
								wait(2)
								Systemutil.Run "C:\Automation\Sym_iEx\BatchFile\Sym_IX_StartService.bat"
								wait(5)
								'********************************
								vStartedSuccessfully =  fnWaitTillExistsStarted(vMsgServiceStarted,vSourceName)
'										************
											If vStartedSuccessfully = True then
												REM *************** 2. Stop the service *********
							'					fnStopService vServiceName,vComputerName
												wait(2)
												Systemutil.Run "C:\Automation\Sym_iEx\BatchFile\Sym_IX_StopService.bat"
												wait(5)
												vStoppedSuccessfully =  fnWaitTillExistsStopped(vMsgServiceStopped,vSourceName)
												If vStoppedSuccessfully = True Then
														Print "Copying the recipient folder to network drive SQL"
														Rem ### Copying recipient files in R:Drive
														SystemUtil.Run "C:\Automation\Sym_iEx\BatchFile\IX_TP11\SQL\IX_TP11_T001_COU_Recipient_SQL.bat"
''																	SystemUtil.Run "C:\Automation\Sym_iEx\BatchFile\Successful_Output_ORA.bat"
												End If
											End If
												
'										***********
print "Test Case TP11_T001 END"
End Function

Function fnSym_TP11_T002_SQL()
								print "Test Case TP11_T002_Starting"
								Systemutil.Run "C:\Automation\Sym_iEx\BatchFile\IX_TP11\Generic\IX_TP11_T002_COU_CSV.bat"
								wait(5)
								Systemutil.Run "C:\Automation\Sym_iEx\BatchFile\IX_TP11\Generic\IX_TP11_T0013_COU_CSV_Local.bat"
								wait(5)
								Print "Import process started"
								fnImportBuild_Sym_TP11
								Print "Import process successfully completed"
								REM - Copy all the import report into network Drive 
								Print "Copy all the import report into network Drive_CopyImportReportSym_TP12_T002_SQL"
								fnCopyImportReportSym_TP11_T002_SQL()
								wait(5)
								Print "Delete all the import report,LO csv files and combined folder"
								fnDeleteFileFolderSQL
								
								
								REM **************************** Second Phase of Service start ,stop and clear  Event Viewer AFTER import ********
								REM ************** 1. Start the service *********************
													Rem - Clear the Event Viewer list ******
								print "SECOND PHASE - CLEAN Event VIEWER SQL"
								Systemutil.Run "C:\Automation\Sym_iEx\BatchFile\ClearEventViewerListApplication.bat"
'										*****************
								REM - Saving Bluelight iExchange manager setting*****
'								fnSaveBLiExMngSetting()
'								
''										**************
								'fnStartService vServiceName,vComputerName
								wait(2)
								Systemutil.Run "C:\Automation\Sym_iEx\BatchFile\Sym_IX_StartService.bat"
								wait(5)
								'********************************
								vStartedSuccessfully =  fnWaitTillExistsStarted(vMsgServiceStarted,vSourceName)
'										************
											If vStartedSuccessfully = True then
												REM *************** 2. Stop the service *********
							'					fnStopService vServiceName,vComputerName
												wait(2)
												Systemutil.Run "C:\Automation\Sym_iEx\BatchFile\Sym_IX_StopService.bat"
												wait(5)
												vStoppedSuccessfully =  fnWaitTillExistsStopped(vMsgServiceStopped,vSourceName)
												If vStoppedSuccessfully = True Then
														Print "Copying the recipient folder to network drive SQL"
														Rem ### Copying recipient files in R:Drive
														SystemUtil.Run "C:\Automation\Sym_iEx\BatchFile\IX_TP11\SQL\IX_TP11_T002_COU_Recipient_SQL.bat"
''																	SystemUtil.Run "C:\Automation\Sym_iEx\BatchFile\Successful_Output_ORA.bat"
												End If
											End If
												
'										***********
print "Test Case TP11_T002 END"
End Function
Function fnSym_TP11_T003_SQL()
								print "Test Case TP11_T003_Starting"
								Systemutil.Run "C:\Automation\Sym_iEx\BatchFile\IX_TP11\Generic\IX_TP11_T003_COU_CSV.bat"
								wait(5)
								Systemutil.Run "C:\Automation\Sym_iEx\BatchFile\IX_TP11\Generic\IX_TP11_T0013_COU_CSV_Local.bat"
								wait(5)
								Print "Import process started"
								fnImportBuild_Sym_TP11
								Print "Import process successfully completed"
								REM - Copy all the import report into network Drive 
								Print "Copy all the import report into network Drive_CopyImportReportSym_TP12_T003_SQL"
								fnCopyImportReportSym_TP11_T003_SQL()
								wait(5)
								Print "Delete all the import report,LO csv files and combined folder"
								fnDeleteFileFolderSQL								
								
								REM **************************** Second Phase of Service start ,stop and clear  Event Viewer AFTER import ********
								REM ************** 1. Start the service *********************
													Rem - Clear the Event Viewer list ******
								print "SECOND PHASE - CLEAN Event VIEWER SQL"
								Systemutil.Run "C:\Automation\Sym_iEx\BatchFile\ClearEventViewerListApplication.bat"
'										*****************
								REM - Saving Bluelight iExchange manager setting*****
'								fnSaveBLiExMngSetting()
'								
''										**************
								'fnStartService vServiceName,vComputerName
								wait(2)
								Systemutil.Run "C:\Automation\Sym_iEx\BatchFile\Sym_IX_StartService.bat"
								wait(5)
								'********************************
								vStartedSuccessfully =  fnWaitTillExistsStarted(vMsgServiceStarted,vSourceName)
'										************
											If vStartedSuccessfully = True then
												REM *************** 2. Stop the service *********
							'					fnStopService vServiceName,vComputerName
												wait(2)
												Systemutil.Run "C:\Automation\Sym_iEx\BatchFile\Sym_IX_StopService.bat"
												wait(5)
												vStoppedSuccessfully =  fnWaitTillExistsStopped(vMsgServiceStopped,vSourceName)
												If vStoppedSuccessfully = True Then
														Print "Copying the recipient folder to network drive SQL"
														Rem ### Copying recipient files in R:Drive
														SystemUtil.Run "C:\Automation\Sym_iEx\BatchFile\IX_TP11\SQL\IX_TP11_T003_COU_Recipient_SQL.bat"
''																	SystemUtil.Run "C:\Automation\Sym_iEx\BatchFile\Successful_Output_ORA.bat"
												End If
											End If
												
'										***********
print "Test Case TP11_T003 END"
End Function
'****************  Sym_TP07 Copying importReport to R:Drive  **********
Function fnCopyImportReportSym_TP07_T002_SQL()
	Systemutil.Run "C:\Automation\Sym_iEx\BatchFile\IX_TP07\SQL\IX_TP07_T002_COU_ImportReport_SQL.bat"
End Function
Function fnCopyImportReportSym_TP07_T003_SQL()
	Systemutil.Run "C:\Automation\Sym_iEx\BatchFile\IX_TP07\SQL\IX_TP07_T003_COU_ImportReport_SQL.bat"
End Function
Function fnCopyImportReportSym_TP07_T004_SQL()
	Systemutil.Run "C:\Automation\Sym_iEx\BatchFile\IX_TP07\SQL\IX_TP07_T004_COU_ImportReport_SQL.bat"
End Function


'****************  Sym_TP09 Copying importReport to R:Drive  **********

Function fnCopyImportReportSym_TP09_T006_SQL()
	Systemutil.Run "C:\Automation\Sym_iEx\BatchFile\IX_TP09\SQL\IX_TP09_T006_COU_ImportReport_SQL.bat"
End Function
Function fnCopyImportReportSym_TP09_T007_SQL()
	Systemutil.Run "C:\Automation\Sym_iEx\BatchFile\IX_TP09\SQL\IX_TP09_T007_COU_ImportReport_SQL.bat"
End Function


'**********************************
'****************  Sym_TP10 Copying importReport to R:Drive  **********
Function fnCopyImportReportSym_TP10_PreTest_SQL()
	Systemutil.Run "C:\Automation\Sym_iEx\BatchFile\IX_TP10\SQL\IX_TP10_PreTest_COU_ImportReport_SQL.bat"
End Function

'*************Sym_TP11 Copying importReport to R:Drive********
Function fnCopyImportReportSym_TP11_T001_SQL()
	Systemutil.Run "C:\Automation\Sym_iEx\BatchFile\IX_TP11\SQL\IX_TP11_T001_COU_ImportReport_SQL.bat"
End Function
Function fnCopyImportReportSym_TP11_T002_SQL()
	Systemutil.Run "C:\Automation\Sym_iEx\BatchFile\IX_TP11\SQL\IX_TP11_T002_COU_ImportReport_SQL.bat"
End Function
Function fnCopyImportReportSym_TP11_T003_SQL()
	Systemutil.Run "C:\Automation\Sym_iEx\BatchFile\IX_TP11\SQL\IX_TP11_T003_COU_ImportReport_SQL.bat"
End Function


Function fnDeleteFileFolderSQL()
'			fnDeleteFileInFolder vCOUFiles
			fnDeleteFolder vCOUCombinedFolder
			fnDeleteFileInFolder vCOUImportErrorFilesSQL
End Function



'*****************************
'Systemutil.Run "C:\Automation\Sym_iEx\BatchFile\Sym_IX_StartService.bat"
'
'vTo = "pradeep.lalung@aligned-assets.co.uk"
'vCC = "pradeep.lalung@ymail.com"
'vTestEnv = "ORACLE"
'fnSend_TestResult_Sym_TP11_SQL vTo,vCC,vTestEnvironment_SQL,vExecutionTime,vSym_iX_Version1,vSym_iX_Service_Date,vSym_iX_Manager_Version1,_
'vSym_iX_Manager_Date,vSym_DB_Version1,vSym_DB_Build_Date,vMigrationCheck_SQL,vCsvFile_SQL,vImportReportFiles_SQL	