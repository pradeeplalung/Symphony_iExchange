'Test Object : Import Additional FRC csv files in Oracle and SQL environment
Option explicit
Dim vCSVPath,vSSFConfigFile,vSSFConfigFileORA,vSSFConfigFileSQL, vSearchImportFolderLine,vReplaceImportFolderLine,vSymOraLogon, vPattern
Dim vBL_App_Path,vBL_App_Setting_Path,vTP_Path ,vTP,vTCD3,vTPS,vTSEC, vUsername,vPassword
Dim WshShell,DeviceReplay,vReport,vReport1,WshShell1,vWelcomePage

Dim vRowControl
'*********** Assigned  Environment variable  values dynamically from XML file *************
Environment.LoadFromFile("C:\Automation\BL_iEx\Xml_File1\BX_TP01_T003.xml")

vBL_App_Path = Environment.Value("vBL_App_Path")	

vBL_App_Setting_Path = Environment.Value("vBL_App_Setting_Path")
vTP_Path = Environment.Value("vTP_Path")					
	
vTP = Environment.Value("vTP") ' Test Plan  Work Sheet ,
    					
vTCD3 =  Environment.Value("vTCD3")

							
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
vCSVPath = Trim(fnReadFromExcel(vTP_Path,vTCD3,2,2))
print "vCSVPath  : "&vCSVPath

'''vReplaceImportFolderLine = "<entry name="&chr(34)&"Import Folder"&chr(34)&">"&vCSVPath&"</entry>"

vRowControl =  fnReadFromExcel(vTP_Path,vTSEC,3,3)


'******************************* Logic *******************
''**************** For both Oracle and SQL environment **************
If  fnReadFromExcel(vTP_Path,vTPS,21,2) = "Yes" and fnReadFromExcel(vTP_Path,vTPS,22,2) = "Yes" Then
		Do while   fnReadFromExcel(vTP_Path,vTCD3,vRowControl,3) <> ""
								If    fnReadFromExcel(vTP_Path,vTCD3,vRowControl,3) = "Yes" and _
									   fnReadFromExcel(vTP_Path,vTSEC,7,5) = "Yes" then

										Select Case vRowControl
										Case 3
											On error resume next
											fnT001 vRowControl,vReplaceImportFolderLine
										Case 4
											On error resume next
											fnT002 vRowControl,vReplaceImportFolderLine  						
										Case else
												print "BX_TP01_T003 - Invalid number Case Number- Test Halts"
										End Select
									End If
									vRowControl = vRowControl+1
		Loop
else
'''****************** For Oracle Environment************************
		        If  fnReadFromExcel(vTP_Path,vTPS,21,2) = "Yes" and fnReadFromExcel(vTP_Path,vTPS,22,2) = "No" and fnReadFromExcel(vTP_Path,vTCD3,3,3) = "Yes" and _
											fnReadFromExcel(vTP_Path,vTSEC,7,5) = "Yes" Then
											On error resume next
											fnT001 vRowControl,vReplaceImportFolderLine
				End If
End if
'***************
'''****************** For SQL Environment************************
If  fnReadFromExcel(vTP_Path,vTPS,21,2) = "No" and fnReadFromExcel(vTP_Path,vTPS,22,2) = "Yes" and  fnReadFromExcel(vTP_Path,vTCD3,4,3) = "Yes" and _
											fnReadFromExcel(vTP_Path,vTSEC,7,5) = "Yes" Then
											On error resume next
											fnT002 vRowControl+1,vReplaceImportFolderLine
End If
'*************************** Copying Report in the network drive *************
SystemUtil.Run "C:\Automation\BL_iEx\Batch Files\T003_FRC_CreateFolderCopyFileIntoIt.bat"
'***************************
Function fnT001(vRowControl,vReplaceImportFolderLine)
					 fnBluelightSettingToORACLE(vBL_App_Setting_Path)
					vReport1 =Trim(fnReadFromExcel(vTP_Path,vTCD3,vRowControl,2)) ' "C:\Automation\Bluelight iExchange\Local_CSVFolder\T003\Report3A_SQL"
					fnDeleteFileInFolder vReport1&"\"					
					'*********************************
					 fnStartIBLImport vBL_App_Path
					''''*********************************
					Wait(10)
					Set vSymOraLogon = SwfWindow("Symphony Oracle Logon")
					If  vSymOraLogon.exist(1) Then
								fnOracleLogOn vUsername, vPassword
								 fnImportProcessToDisplayReport3 vCSVPath
								fnReportSaveSQL234(vReport1)
					Else 
								 fnImportProcessToDisplayReport3 vCSVPath
								 fnReportSaveSQL234(vReport1)
					End if
End Function
'''*************************************
Function fnT002(vRowControl,vReplaceImportFolderLine)
					 fnBluelightSettingToSQL(vBL_App_Setting_Path)
					vReport1 =Trim(fnReadFromExcel(vTP_Path,vTCD3,vRowControl,2)) ' "C:\Automation\Bluelight iExchange\Local_CSVFolder\T003\Report3A_SQL"
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
								 fnImportProcessToDisplayReport3 vCSVPath
								fnReportSaveSQL234(vReport1)
					Else 
								 fnImportProcessToDisplayReport3 vCSVPath
								 fnReportSaveSQL234(vReport1)
					End if
End Function
'***************************
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
Function fnImportProcessToDisplayReport1()
'				SwfWindow("Symphony Import").Activate
				SwfWindow("Symphony Import").SwfObject("OS Insight - AddressBase").Click 51,33
				Wait(3)
				SwfWindow("Symphony Import").SwfObject("Next >").Click 36,11
				SwfWindow("Symphony Import").SwfObject("Next >").Click 36,11
				SwfWindow("Symphony Import").SwfObject("Next >").Click 36,11
				SwfWindow("Symphony Import").SwfObject("Next >").Click 36,11
'				*************** Logic when csv file import completed then fnWaitTilExist will be true ****
				If  fnWaitTilExists = True then
					SwfWindow("Symphony Import").SwfObject("Display Report").Click 45,11
				end if
'				SwfWindow("Symphony Import").Close
End Function
'*****************
Function fnImportProcessToDisplayReportOld12(vCSVPathFile)
				SwfWindow("Symphony Import").SwfObject("OS Insight - AddressBase").Click 132,61'For Local
'''''				 SwfWindow("Symphony Import").SwfObject("OS Insight - AddressBase").Click 23,81' Additional Data
				Wait(3)
				SwfWindow("Symphony Import").SwfObject("Next >").Click 36,11
				SwfWindow("Symphony Import").SwfObject("Next >").Click 36,11
				SwfWindow("Symphony Import").SwfObject("Next >").Click 36,11
				SwfWindow("Symphony Import").SwfObject("Next >").Click 36,11
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
'				*************** Logic when csv file import completed then fnWaitTilExist will be true ****
				If  fnWaitTilExists = True then
					SwfWindow("Symphony Import").SwfObject("Display Report").Click 45,11
				end if
'				SwfWindow("Symphony Import").Close
End Function
'******************
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
Function fnReportSaveSQL1(vReport) 
			'vReport = "C:\Automation\BL_iEx\CSV\Import Files\COU_Report\ORACLE"
'			vReport = "C:\Automation\BL_iEx\CSV\T004_Report\SQL"
			Set DeviceReplay = CreateObject("Mercury.DeviceReplay")
			Set WshShell1 = CreateObject("WScript.Shell")
			SwfWindow("Symphony Import").SwfWindow("Import Report").SwfObject("Save Error File").Click 59,15
'			DeviceReplay.SendString(vReportTrick)	' Type the path where report to be stored 
'			SwfWindow("Symphony Import").SwfWindow("Import Report").Dialog("Export Import Errors As").WinButton("Save").Click ' set the report path
			Wait(5)
			DeviceReplay.SendString vReport&"\Import Errors.txt" 	' Type the path where report to be stored 
			Print "Hello"
			SwfWindow("Symphony Import").SwfWindow("Import Report").Dialog("Export Import Errors As").WinButton("Save").Click ' set the report path
'			Wait(3)
Print "Hello2"
'			******** Saving text report ********
'			SwfWindow("Symphony Import").SwfWindow("Import Report").Dialog("Export Import Errors As").WinButton("Save").Click	
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
'*******************
Function fnImportProcessToDisplayReport2(vCSVPathFile)
				SwfWindow("Symphony Import").SwfObject("OS Insight - AddressBase").Click 132,61'For Local
'''''				 SwfWindow("Symphony Import").SwfObject("OS Insight - AddressBase").Click 23,81' Additional Data
				Wait(3)
				SwfWindow("Symphony Import").SwfObject("Next >").Click 36,11
				SwfWindow("Symphony Import").SwfObject("Next >").Click 36,11
				SwfWindow("Symphony Import").SwfObject("Next >").Click 36,11
				SwfWindow("Symphony Import").SwfObject("Next >").Click 36,11
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
'				*************** Logic when csv file import completed then fnWaitTilExist will be true ****
				If  fnWaitTilExists = True then
					SwfWindow("Symphony Import").SwfObject("Display Report").Click 45,11
				end if
'				SwfWindow("Symphony Import").Close
End Function
'******************
Function fnImportProcessToDisplayReport3(vCSVPathFile)
'				SwfWindow("Symphony Import").Activate
''''				SwfWindow("Symphony Import").SwfObject("OS Insight - AddressBase").Click 51,33 ' For COU 
''				SwfWindow("Symphony Import").SwfObject("OS Insight - AddressBase").Click 132,61'For Local
				' In case Default screen appears with 'NEXT' button after fresh installation
				If Trim(OptionalStep.SwfWindow("Symphony Import").SwfObject("Welcome to the Symphony").GetROProperty("Text")) = vWelcomePage Then
				
						SwfWindow("Symphony Import").SwfObject("Next >").Click 36,11
						print "Welcome page found in FRC_LO"
				End If
				 SwfWindow("Symphony Import").SwfObject("OS Insight - AddressBase").Click 23,81' Additional Data
				Wait(3)
				SwfWindow("Symphony Import").SwfObject("Next >").Click 36,11
				SwfWindow("Symphony Import").SwfObject("Next >").Click 36,11
				SwfWindow("Symphony Import").SwfObject("Next >").Click 36,11
'				SwfWindow("Symphony Import").SwfObject("Next >").Click 36,11
'''**************************** Selecting the LO cvs file 
				SwfWindow("Symphony Import").SwfObject("txtFilename").Click 383,13
''				SwfWindow("Symphony Import").SwfEdit("SwfEdit").Type vCSVPath
'				SwfWindow("Symphony Import").Dialog("Import").WinButton("OK").Click
				wait(2)
				SwfWindow("Symphony Import").Dialog("Select NLPG CSV Import File").Activate
				SwfWindow("Symphony Import").Dialog("Select NLPG CSV Import File").WinEdit("File name:").Set vCSVPathFile
				SwfWindow("Symphony Import").Dialog("Select NLPG CSV Import File").WinButton("Open").Click

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
'****************
Function fnImportProcessToDisplayReport4(vCSVPathFile)
'				SwfWindow("Symphony Import").Activate
''''				SwfWindow("Symphony Import").SwfObject("OS Insight - AddressBase").Click 51,33 ' For COU 
''				SwfWindow("Symphony Import").SwfObject("OS Insight - AddressBase").Click 132,61'For Local
				 SwfWindow("Symphony Import").SwfObject("OS Insight - AddressBase").Click 23,81' Additional Data
				Wait(3)
				SwfWindow("Symphony Import").SwfObject("Next >").Click 36,11
				SwfWindow("Symphony Import").SwfObject("Next >").Click 36,11
				SwfWindow("Symphony Import").SwfObject("Next >").Click 36,11
'				SwfWindow("Symphony Import").SwfObject("Next >").Click 36,11
'''**************************** Selecting the LO cvs file 
				SwfWindow("Symphony Import").SwfObject("txtFilename").Click 383,13
''				SwfWindow("Symphony Import").SwfEdit("SwfEdit").Type vCSVPath
'				SwfWindow("Symphony Import").Dialog("Import").WinButton("OK").Click
				wait(2)
				SwfWindow("Symphony Import").Dialog("Select NLPG CSV Import File").Activate
				SwfWindow("Symphony Import").Dialog("Select NLPG CSV Import File").WinEdit("File name:").Set vCSVPathFile
				SwfWindow("Symphony Import").Dialog("Select NLPG CSV Import File").WinButton("Open").Click

'''''****************************
				wait(2)
				SwfWindow("Symphony Import").SwfObject("Next >").Click 36,11
				SwfWindow("Symphony Import").SwfObject("Next >").Click 36,11
'				*************** Logic when csv file import completed then fnWaitTilExist will be true ****
				If  fnWaitTilExists = True then
					SwfWindow("Symphony Import").SwfObject("Display Report").Click 45,11
				end if
'				SwfWindow("Symphony Import").Close
End Function
'****************
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
'******************