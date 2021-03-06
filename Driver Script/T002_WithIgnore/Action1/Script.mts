﻿'Test Objective : Import LO csv file in Oracle and SQL environment
Option explicit
Dim vCSVPath,vSSFConfigFile,vSSFConfigFileORA,vSSFConfigFileSQL, vSearchImportFolderLine,vReplaceImportFolderLine,vSymOraLogon, vPattern
Dim vBL_App_Path,vBL_App_Setting_Path,vTP_Path ,vTP,vTCD2,vTPS,vTSEC, vUsername,vPassword
Dim WshShell,DeviceReplay,vReport,vReport1,WshShell1,vWelcomePage

Dim vRowControl
'*********** Assigned  Environment variable  values dynamically from XML file *************
Environment.LoadFromFile("C:\Automation\BL_iEx\Xml_File\BX_TP01_T002.xml")

vBL_App_Path = Environment.Value("vBL_App_Path")	

vBL_App_Setting_Path = Environment.Value("vBL_App_Setting_Path")
vTP_Path = Environment.Value("vTP_Path")					
	
vTP = Environment.Value("vTP") ' Test Plan  Work Sheet ,
    					
vTCD2 =  Environment.Value("vTCD2")

							
vTPS = Environment.Value("vTPS")


vTSEC = Environment.Value("vTSEC")
vWelcomePage = Environment.Value("vWelcomePage")


vUsername = Environment.Value("vUsername")
print vUsername

vPassword = Environment.Value("vPassword")
print vPassword
'**************************

vSSFConfigFile =Trim(fnReadFromExcel(vTP_Path,vTPS,6,2))   '"C:\SSF\BLUELIGHT_NAG.SSF"
vSSFConfigFile =Split(vSSFConfigFile,";")
vSSFConfigFileORA = Trim(vSSFConfigFile(0))
vSSFConfigFileSQL = Trim(vSSFConfigFile(1))

''vPattern = ".*Import Folder.*"
vCSVPath = Trim(fnReadFromExcel(vTP_Path,vTCD2,2,2))
print "vCSVPath  : "&vCSVPath

'''vReplaceImportFolderLine = "<entry name="&chr(34)&"Import Folder"&chr(34)&">"&vCSVPath&"</entry>"

vRowControl =  fnReadFromExcel(vTP_Path,vTSEC,3,3)


'******************************* Logic *******************
''**************** For both Oracle and SQL environment **************
If  fnReadFromExcel(vTP_Path,vTPS,21,2) = "Yes" and fnReadFromExcel(vTP_Path,vTPS,22,2) = "Yes" Then
		Do while   fnReadFromExcel(vTP_Path,vTCD2,vRowControl,3) <> ""
								If    fnReadFromExcel(vTP_Path,vTCD2,vRowControl,3) = "Yes" and _
									   fnReadFromExcel(vTP_Path,vTSEC,5,5) = "Yes" then

										Select Case vRowControl
										Case 3
											On error resume next
											fnT001 vRowControl,vReplaceImportFolderLine
										Case 4
											On error resume next
											fnT002 vRowControl,vReplaceImportFolderLine  						
										Case else
												print "BX_TP01_T002 - Invalid number Case Number- Test Halts"
										End Select
									End If
									vRowControl = vRowControl+1
		Loop
else
'''****************** For Oracle Environment************************
		        If  fnReadFromExcel(vTP_Path,vTPS,21,2) = "Yes" and fnReadFromExcel(vTP_Path,vTPS,22,2) = "No" and fnReadFromExcel(vTP_Path,vTCD2,3,3) = "Yes" and _
											fnReadFromExcel(vTP_Path,vTSEC,5,5) = "Yes" Then
											On error resume next
											fnT001 vRowControl,vReplaceImportFolderLine
				End If
End if
'***************
'''****************** For SQL Environment************************
If  fnReadFromExcel(vTP_Path,vTPS,21,2) = "No" and fnReadFromExcel(vTP_Path,vTPS,22,2) = "Yes" and  fnReadFromExcel(vTP_Path,vTCD2,4,3) = "Yes" and _
											fnReadFromExcel(vTP_Path,vTSEC,5,5) = "Yes" Then
											On error resume next
											fnT002 vRowControl+1,vReplaceImportFolderLine
End If
'***************************
SystemUtil.Run "C:\Automation\BL_iEx\Batch Files\T002_LO_CreateFolderCopyFileIntoIt.bat"
'******************************
Function fnT001(vRowControl,vReplaceImportFolderLine)
					 fnBluelightSettingToORACLE(vBL_App_Setting_Path)
					vReport1 =Trim(fnReadFromExcel(vTP_Path,vTCD2,vRowControl,2)) ' "C:\Automation\Bluelight iExchange\Local_CSVFolder\T003\Report3A_SQL"
					fnDeleteFileInFolder vReport1&"\"					
					'*********************************
					 fnStartIBLImport vBL_App_Path
					''''*********************************
					Wait(10)
					Set vSymOraLogon = SwfWindow("Symphony Oracle Logon")
					If  vSymOraLogon.exist(1) Then
								fnOracleLogOn vUsername, vPassword
								 fnImportProcessToDisplayReport2 vCSVPath
								fnReportSaveSQL234(vReport1)
					Else 
								 fnImportProcessToDisplayReport2 vCSVPath
								 fnReportSaveSQL234(vReport1)
					End if
End Function
'''*************************************
Function fnT002(vRowControl,vReplaceImportFolderLine)
					 fnBluelightSettingToSQL(vBL_App_Setting_Path)
					vReport1 =Trim(fnReadFromExcel(vTP_Path,vTCD2,vRowControl,2)) ' "C:\Automation\Bluelight iExchange\Local_CSVFolder\T003\Report3A_SQL"
					fnDeleteFileInFolder vReport1&"\"
'					vSearchImportFolderLine =  fnSearchLineOfPatternInFile(vSSFConfigFile,vPattern)
'					fnReplaceLineWithSpecificValue vSSFConfigFile,vSearchImportFolderLine,vReplaceImportFolderLine,vSSFConfigFile
					
''''					'*********************************
					 fnStartIBLImport vBL_App_Path
''''					''''*********************************
					Wait(10)
					Set vSymOraLogon = SwfWindow("Symphony Oracle Logon")
					If  vSymOraLogon.exist(1) Then
								fnOracleLogOn vUsername, vPassword
								 fnImportProcessToDisplayReport2 vCSVPath
								fnReportSaveSQL234(vReport1)
					Else 
								 fnImportProcessToDisplayReport2 vCSVPath
								 fnReportSaveSQL234(vReport1)
					End if
''''************************** 
End Function
'***************************
'Function fnSwitchDataBaseSQL(vBL_App_Setting_Path)
'   SystemUtil.Run vBL_App_Setting_Path
'''''*********************************
'					Wait(10)
'					Set vSymOraLogon = SwfWindow("Symphony Oracle Logon")
'					If  vSymOraLogon.exist(1) Then
'										fnOracleLogOn vUsername, vPassword
'										
'										SwfWindow("Symphony Settings").SwfObject("Oracle").Click 27,7 @@ hightlight id_;_395008_;_script infofile_;_ZIP::ssf45.xml_;_
'		
'										SwfWindow("Symphony Settings").SwfObject("Test Connection").Click 53,12 @@ hightlight id_;_197898_;_script infofile_;_ZIP::ssf50.xml_;_
'										Wait(2)							
'										SwfWindow("Symphony Settings").Dialog("Symphony Settings").WinButton("OK").Click @@ hightlight id_;_67418_;_script infofile_;_ZIP::ssf46.xml_;_
'										Wait(2)		
'										SwfWindow("Symphony Settings").SwfObject("OK").Click 45,11
'										Wait(2)		
'										SwfWindow("Symphony Settings").Dialog("Symphony Settings").WinButton("OK").Click @@ hightlight id_;_1050332_;_script infofile_;_ZIP::ssf51.xml_;_
'					else @@ hightlight id_;_2360276_;_script infofile_;_ZIP::ssf47.xml_;_
'										SwfWindow("Symphony Settings").SwfObject("Oracle").Click 164,8 @@ hightlight id_;_2295058_;_script infofile_;_ZIP::ssf53.xml_;_
'										SwfWindow("Symphony Settings").SwfObject("Test Connection").Click 32,15 @@ hightlight id_;_2098516_;_script infofile_;_ZIP::ssf54.xml_;_
'										wait(5)
'										fnOracleLogOnWithin vUsername, vPassword
'										Wait(3)	
'''										*************************
''										SwfWindow("Symphony Settings").SwfWindow("Symphony Oracle Logon").SwfEdit("SwfEdit").Set "AAUSER" @@ hightlight id_;_330004_;_script infofile_;_ZIP::ssf55.xml_;_
''										SwfWindow("Symphony Settings").SwfWindow("Symphony Oracle Logon").SwfObject("utePassword").Click 26,6 @@ hightlight id_;_198950_;_script infofile_;_ZIP::ssf56.xml_;_
''										SwfWindow("Symphony Settings").SwfWindow("Symphony Oracle Logon").SwfEdit("SwfEdit").SetSecure "544921e1a930dcd4cd34fc4b2ddc174a1bae" @@ hightlight id_;_1247540_;_script infofile_;_ZIP::ssf57.xml_;_
''										SwfWindow("Symphony Settings").SwfWindow("Symphony Oracle Logon").SwfObject("OK").Click 53,15 @@ hightlight id_;_198952_;_script infofile_;_ZIP::ssf58.xml_;_
'''										*************************
'										SwfWindow("Symphony Settings").Dialog("Symphony Settings").WinButton("OK").Click @@ hightlight id_;_264466_;_script infofile_;_ZIP::ssf59.xml_;_
'										Wait(3)	
'										SwfWindow("Symphony Settings").SwfObject("OK").Click 44,8 @@ hightlight id_;_2360882_;_script infofile_;_ZIP::ssf60.xml_;_
'										Wait(3)	
'										SwfWindow("Symphony Settings").Dialog("Symphony Settings").WinButton("OK").Click @@ hightlight id_;_985364_;_script infofile_;_ZIP::ssf61.xml_;_
'							
'					End if
'End Function
''''*****************************

Function fnBluelightSettingToSQL(vBL_App_Setting_Path)
		   SystemUtil.Run vBL_App_Setting_Path
		   ''''*********************************

		Wait(10)
		Set vSymOraLogon = SwfWindow("Symphony Oracle Logon")
		If  vSymOraLogon.exist(1) Then
			fnOracleLogOn vUsername, vPassword
			SwfWindow("Symphony Settings").SwfObject("C:\SSF\BLUELIGHT_NAG.SSF").Click 714,14
			SwfWindow("Symphony Settings").Dialog("Select a settings file").WinEdit("File name:").Set vSSFConfigFileSQL
			SwfWindow("Symphony Settings").Dialog("Select a settings file").WinButton("Open").Click
			SwfWindow("Symphony Settings").SwfObject("Test Connection").Click 56,11
			SwfWindow("Symphony Settings").Dialog("Symphony Settings").WinButton("OK").Click
			SwfWindow("Symphony Settings").SwfObject("OK").Click 32,11
			SwfWindow("Symphony Settings").Dialog("Symphony Settings").WinButton("Yes").Click
		Else 
			OptionalStep.SwfWindow("Symphony Settings").SwfObject("OK").Click 32,11
		End if
End Function
'*************
Function fnBluelightSettingToORACLE(vBL_App_Setting_Path)
					SystemUtil.Run vBL_App_Setting_Path		'''*********************************       
					Wait(10)  
					If  SwfWindow("Symphony Oracle Logon").exist(1) Then
						fnOracleLogOn vUsername, vPassword
						SwfWindow("Symphony Settings").SwfObject("OK").Click 32,11
				   Else
						SwfWindow("Symphony Settings").SwfObject("C:\SSF\BLUELIGHT_NAG.SSF").Click 714,14
						print "vSSFConfigFileORA   :"&vSSFConfigFileORA
						SwfWindow("Symphony Settings").Dialog("Select a settings file").WinEdit("File name:").Set vSSFConfigFileORA
						SwfWindow("Symphony Settings").Dialog("Select a settings file").WinButton("Open").Click
						fnSymOraLogonWithinSetting vUsername, vPassword 		
						SwfWindow("Symphony Settings").SwfObject("Test Connection").Click 56,11
						SwfWindow("Symphony Settings").Dialog("Symphony Settings").WinButton("OK").Click
						SwfWindow("Symphony Settings").SwfObject("OK").Click 32,11
						SwfWindow("Symphony Settings").Dialog("Symphony Settings").WinButton("Yes").Click
					End if
		   
End Function
'''**************************************
Function fnSymOraLogonWithinSetting(vUsername, vPassword)
   			Set WshShell = CreateObject("Wscript.Shell")
            Set DeviceReplay = CreateObject("Mercury.DeviceReplay")
			DeviceReplay.SendString(vUsername)
			WshShell.SendKeys "{TAB}"
			DeviceReplay.SendString(vPassword)
			SwfWindow("Symphony Settings").SwfWindow("Symphony Oracle Logon").SwfObject("OK").Click 32,13
End Function
'*******************  Start the import  process *******************************************
Function fnStartIBLImport(vApp_Path)
				SystemUtil.Run vApp_Path
End Function
'**********  For oracle login window  to Open Bluelight import setting application*****
Function fnOracleLogOn(vUsername,vPassword)
			Set WshShell = CreateObject("Wscript.Shell")
            Set DeviceReplay = CreateObject("Mercury.DeviceReplay")
			DeviceReplay.SendString(vUsername)
			WshShell.SendKeys "{TAB}"
			DeviceReplay.SendString(vPassword)
			SwfWindow("Symphony Oracle Logon").SwfObject("OK").Click()
End Function
'**************or oracle login window  to  switch to Oracle database within Bluelight import setting application*****
Function fnOracleLogOnWithin(vUsername,vPassword)
			Set WshShell = CreateObject("Wscript.Shell")
            Set DeviceReplay = CreateObject("Mercury.DeviceReplay")
			DeviceReplay.SendString(vUsername)
			WshShell.SendKeys "{TAB}"
			DeviceReplay.SendString(vPassword)
        	SwfWindow("Symphony Settings").SwfWindow("Symphony Oracle Logon").SwfObject("OK").Click 15,14
End Function
'************************** 
Function fnImportProcessToDisplayReport2(vCSVPathFile)
				' In case Default screen appears with 'NEXT' button after fresh installation
'				OptionalStep.SwfWindow("Symphony Import").SwfObject("Next >").Click 36,11
Print "Helloooooooo"
				If Trim(OptionalStep.SwfWindow("Symphony Import").SwfObject("Welcome to the Symphony").GetROProperty("Text")) = vWelcomePage Then
				
						SwfWindow("Symphony Import").SwfObject("Next >").Click 36,11
						print "Welcome page found in Add_LO2"
				End If
				SwfWindow("Symphony Import").SwfObject("OS Insight - AddressBase").Click 132,61'For Local
'''''				 SwfWindow("Symphony Import").SwfObject("OS Insight - AddressBase").Click 23,81' Additional Data
				Wait(3)
'				********** If IgnorFailed Data and Reprocess screen appear ***
				if (instr(SwfWindow("Symphony Import").SwfObject("IgnoreFailedData").GetROProperty("Text"), "last") <> 0) then 
				
						wait(3)
						SwfWindow("Symphony Import").SwfObject("grpProcess").Click 272,38
						wait(2)
						SwfWindow("Symphony Import").SwfObject("Next >").Click 26,12
						print "Got ignore"
				end if	

				SwfWindow("Symphony Import").SwfObject("Next >").Click 36,11
				SwfWindow("Symphony Import").SwfObject("Next >").Click 36,11
				SwfWindow("Symphony Import").SwfObject("Next >").Click 36,11
'				SwfWindow("Symphony Import").SwfObject("Next >").Click 36,11
'''**************************** Selecting the LO cvs file 
				SwfWindow("Symphony Import").SwfObject("txtFilename").Click 383,13
''				SwfWindow("Symphony Import").SwfEdit("SwfEdit").Type vCSVPath
'				SwfWindow("Symphony Import").Dialog("Import").WinButton("OK").Click
				wait(2)
'				SwfWindow("Symphony Import").Dialog("Select NLPG CSV Import File").Activate
				SwfWindow("Symphony Import").Dialog("Select NLPG CSV Import").WinEdit("File name:").Set vCSVPathFile
				SwfWindow("Symphony Import").Dialog("Select NLPG CSV Import").WinButton("Open").Click
				

'''''****************************
				wait(2)
				SwfWindow("Symphony Import").SwfObject("Next >").Click 36,11
				SwfWindow("Symphony Import").SwfObject("Next >").Click 36,11
				OptionalStep.SwfWindow("Symphony Import").Dialog("Import").WinButton("OK").click
				OptionalStep.Window("Symphony Import").Dialog("Import").WinButton("OK").click
'				*************** Logic when csv file import completed then fnWaitTilExist will be true ****
				If  fnWaitTilExists = True then
					SwfWindow("Symphony Import").SwfObject("Display Report").Click 45,11
				end if
'				SwfWindow("Symphony Import").Close
End Function
'************************
Function fnImportProcessToDisplayReportOLD (vCSVPath)
'				SwfWindow("Symphony Import").Activate
''''				SwfWindow("Symphony Import").SwfObject("OS Insight - AddressBase").Click 51,33 ' For COU 
				SwfWindow("Symphony Import").SwfObject("OS Insight - AddressBase").Click 132,61'For Local
				Wait(3)
				SwfWindow("Symphony Import").SwfObject("Next >").Click 36,11
				SwfWindow("Symphony Import").SwfObject("Next >").Click 36,11
				SwfWindow("Symphony Import").SwfObject("Next >").Click 36,11
'				SwfWindow("Symphony Import").SwfObject("Next >").Click 36,11
'''**************************** Selecting the LO cvs file 
'				SwfWindow("Symphony Import").SwfObject("txtFilename").Click 383,13
				SwfWindow("Symphony Import").SwfEdit("SwfEdit").Set vCSVPath
'				SwfWindow("Symphony Import").Dialog("Import").WinButton("OK").Click
'				SwfWindow("Symphony Import").Dialog("Select NAG CSV Import").WinEdit("File name:").Type  vCSVPath ''"C:\Automation\BL_iEx\CSV\Import Files\LO\ADDRESS_BASE_LO_PROP_2014-10-01.csv"
'				SwfWindow("Symphony Import").Dialog("Select NAG CSV Import").WinEdit("File name:").Type  micReturn 
'''''****************************
				SwfWindow("Symphony Import").SwfObject("Next >").Click 36,11
				SwfWindow("Symphony Import").SwfObject("Next >").Click 36,11
'				*************** Logic when csv file import completed then fnWaitTilExist will be true ****
				If  fnWaitTilExists = True then
					SwfWindow("Symphony Import").SwfObject("Display Report").Click 45,11
				end if
'				SwfWindow("Symphony Import").Close
End Function
''************************

Function fnImportProcessToDisplayReport123()
'				SwfWindow("Symphony Import").Activate
'				SwfWindow("Symphony Import").SwfObject("OS Insight - AddressBase").Click 51,33 ' For COU
				SwfWindow("Symphony Import").SwfObject("OS Insight - AddressBase").Click 132,61'For Local
				Wait(3)
				SwfWindow("Symphony Import").SwfObject("Next >").Click 36,11
				SwfWindow("Symphony Import").SwfObject("Next >").Click 36,11
				SwfWindow("Symphony Import").SwfObject("uceBulkImport").Click 5,12
                SwfWindow("Symphony Import").SwfObject("Next >").Click 36,11
				SwfWindow("Symphony Import").SwfObject("Next >").Click 36,11
'				*************** Logic when csv file import completed then fnWaitTilExist will be true ****
				If  fnWaitTilExists = True then
					SwfWindow("Symphony Import").SwfObject("Display Report").Click 45,11
				end if
'				SwfWindow("Symphony Import").Close
End Function
'*****************
Function fnReportSave(vReport) 
'			vReport = "C:\Automation\Bluelight iExchange\Local_CSVFolder\T002\Report"
			Set DeviceReplay = CreateObject("Mercury.DeviceReplay")
			Set WshShell1 = CreateObject("WScript.Shell")
			SwfWindow("Symphony Import").SwfWindow("Import Report").SwfObject("Save Error File").Click 59,15
			DeviceReplay.SendString(vReport)	' Type the path where report to be stored 
			SwfWindow("Symphony Import").SwfWindow("Import Report").Dialog("Export Import Errors As").WinButton("Save").Click ' set the report path
'			******** Saving text report ********
			SwfWindow("Symphony Import").SwfWindow("Import Report").Dialog("Export Import Errors As").WinButton("Save").Click	
'			******** Saving excel Report *****
			SwfWindow("Symphony Import").SwfWindow("Import Report").SwfObject("Export Report").Click 58,11
			SwfWindow("Symphony Import").SwfWindow("Import Report").Dialog("Export Import Report As").WinButton("Save").Click
'			*********** Saving pdf report ******
			SwfWindow("Symphony Import").SwfWindow("Import Report").SwfObject("Save Report").Click 50,13
			SwfWindow("Symphony Import").SwfWindow("Import Report").Dialog("Save Import Report As").WinButton("Save").Click
			SwfWindow("Symphony Import").SwfWindow("Import Report").Dialog("Save Report").WinButton("OK").Click
			SwfWindow("Symphony Import").SwfWindow("Import Report").SwfObject("Exit").Click 41,12
End Function
'**********************************
Function fnReportSaveSQL234(vReport) 
'			vReportTrick = "C:\Automation\BL_iEx\CSV\TestTrick"
'			vReport = "C:\Automation\BL_iEx\CSV\T004_Report\SQL"
			Set DeviceReplay = CreateObject("Mercury.DeviceReplay")
			Set WshShell1 = CreateObject("WScript.Shell")
			SwfWindow("Symphony Import").SwfWindow("Import Report").SwfObject("Save Error File").Click 59,15
'			DeviceReplay.SendString(vReportTrick)	' Type the path where report to be stored 
'			SwfWindow("Symphony Import").SwfWindow("Import Report").Dialog("Export Import Errors As").WinButton("Save").Click ' set the report path
			Wait(5)
			DeviceReplay.SendString vReport 	' Type the path where report to be stored 
			Print "Hello"
			SwfWindow("Symphony Import").SwfWindow("Import Report").Dialog("Export Import Errors As").WinButton("Save").Click ' set the report path
'			Wait(3)
Print "Hello2"
'			******** Saving text report ********
			SwfWindow("Symphony Import").SwfWindow("Import Report").Dialog("Export Import Errors As").WinButton("Save").Click	
'			******** Saving excel Report *****
			Wait(3)
			SwfWindow("Symphony Import").SwfWindow("Import Report").SwfObject("Export Report").Click 58,11
			Wait(3)
			SwfWindow("Symphony Import").SwfWindow("Import Report").Dialog("Export Import Report As").WinButton("Save").Click
'			*********** Saving pdf report ******
			SwfWindow("Symphony Import").SwfWindow("Import Report").SwfObject("Save Report").Click 50,13
			SwfWindow("Symphony Import").SwfWindow("Import Report").Dialog("Save Import Report As").WinButton("Save").Click
			SwfWindow("Symphony Import").SwfWindow("Import Report").Dialog("Save Report").WinButton("OK").Click
			SwfWindow("Symphony Import").SwfWindow("Import Report").SwfObject("Exit").Click 41,12
			Set DeviceReplay = Nothing
			Set WshShell1 = Nothing
End Function
'*****************