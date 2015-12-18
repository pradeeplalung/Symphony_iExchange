'Test Object : Import COU csv file in Oracle and SQL environment

'Option explicit
Dim vCSVPath,vSSFConfigFile,vSSFConfigFileORA,vSSFConfigFileSQL, vSearchImportFolderLine,vReplaceImportFolderLine,vSymOraLogon, vPattern
Dim vBL_App_Path,vBL_App_Setting_Path,vTP_Path ,vTP,vTCD1,vTPS,vTSEC, vUsername,vPassword
Dim WshShell,DeviceReplay,vReport,vReport1,WshShell1

Dim vRowControl
'*********** Assigned  Environment variable  values dynamically from XML file *************
Environment.LoadFromFile("C:\Automation\BL_iEx\Xml_File\BX_TP01_T001.xml")

vBL_App_Path = Environment.Value("vBL_App_Path")	

vBL_App_Setting_Path = Environment.Value("vBL_App_Setting_Path")
vTP_Path = Environment.Value("vTP_Path")					
	
vTP = Environment.Value("vTP") ' Test Plan  Work Sheet ,
    					
vTCD1 =  Environment.Value("vTCD1")

							
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
print "vSSFConfigFileORA  :" &vSSFConfigFileORA
print "vSSFConfigFileSQL  :" &vSSFConfigFileSQL
vPattern = ".*Import Folder.*"
vCSVPath = Trim(fnReadFromExcel(vTP_Path,vTCD1,2,2))
print "vCSVPath  :"&vCSVPath

vReplaceImportFolderLine = "<entry name="&chr(34)&"Import Folder"&chr(34)&">"&vCSVPath&"</entry>"
print "vReplaceImportFolderLine  :"&vReplaceImportFolderLine

vRowControl =  fnReadFromExcel(vTP_Path,vTSEC,3,3)


'******************************* Logic *******************
''**************** For both Oracle and SQL environment **************
If  fnReadFromExcel(vTP_Path,vTPS,21,2) = "Yes" and fnReadFromExcel(vTP_Path,vTPS,22,2) = "Yes" Then
		Do while   fnReadFromExcel(vTP_Path,vTCD1,vRowControl,3) <> ""
								If    fnReadFromExcel(vTP_Path,vTCD1,vRowControl,3) = "Yes" and _
									   fnReadFromExcel(vTP_Path,vTSEC,3,5) = "Yes" then

										Select Case vRowControl
										Case 3
											On error resume next
											fnT001 vRowControl,vReplaceImportFolderLine
										Case 4
											On error resume next
											fnT002 vRowControl,vReplaceImportFolderLine  						
										Case else
												print "BX_TP01_T001 - Invalid number Case Number- Test Halts"
										End Select
									End If
									vRowControl = vRowControl+1
		Loop
else
'''****************** For Oracle Environment************************
		        If  fnReadFromExcel(vTP_Path,vTPS,21,2) = "Yes" and fnReadFromExcel(vTP_Path,vTPS,22,2) = "No" and fnReadFromExcel(vTP_Path,vTCD1,3,3) = "Yes" and _
											fnReadFromExcel(vTP_Path,vTSEC,3,5) = "Yes" Then
											On error resume next
											fnT001 vRowControl,vReplaceImportFolderLine
				End If
End If
'''****************** For SQL Environment************************
If  fnReadFromExcel(vTP_Path,vTPS,21,2) = "No" and fnReadFromExcel(vTP_Path,vTPS,22,2) = "Yes" and  fnReadFromExcel(vTP_Path,vTCD1,4,3) = "Yes" and _
											fnReadFromExcel(vTP_Path,vTSEC,3,5) = "Yes" Then
											On error resume next
											fnT002 vRowControl+1,vReplaceImportFolderLine
End If

'***************
Function fnT001(vRowControl,vReplaceImportFolderLine)
					 fnBluelightSettingToORACLE(vBL_App_Setting_Path)
					vReport1 =Trim(fnReadFromExcel(vTP_Path,vTCD1,vRowControl,2)) ' "C:\Automation\BL_iEx\CSV\Import Files\COU_Report\ORACLE"
					'msgbox vReport1
					fnDeleteFileInFolder vReport1&"\"
					vSearchImportFolderLine =  fnSearchLineOfPatternInFile(vSSFConfigFileORA,vPattern)
					fnReplaceLineWithSpecificValue vSSFConfigFileORA,vSearchImportFolderLine,vReplaceImportFolderLine,vSSFConfigFileORA
					
					'*********************************
					 fnStartIBLImport vBL_App_Path
					''''*********************************
					Wait(10)
					Set vSymOraLogon = SwfWindow("Symphony Oracle Logon")
					If  vSymOraLogon.exist(1) Then
								fnOracleLogOn vUsername, vPassword
								 fnImportProcessToDisplayReport1
								fnReportSaveSQL1(vReport1)

					End if
'************************** 
End Function
Function fnT002(vRowControl,vReplaceImportFolderLine)
					  fnBluelightSettingToSQL(vBL_App_Setting_Path)
					vReport1 =Trim(fnReadFromExcel(vTP_Path,vTCD1,vRowControl,2)) '  "C:\Automation\BL_iEx\CSV\Import Files\COU_Report\SQL"
					fnDeleteFileInFolder vReport1&"\"
					vSearchImportFolderLine =  fnSearchLineOfPatternInFile(vSSFConfigFileSQL,vPattern)
					print "vSearchImportFolderLine  :"&vSearchImportFolderLine
					fnReplaceLineWithSpecificValue vSSFConfigFileSQL,vSearchImportFolderLine,vReplaceImportFolderLine,vSSFConfigFileSQL

					
''''					'*********************************
					 fnStartIBLImport vBL_App_Path
''''					''''*********************************
					Wait(10)
					Set vSymOraLogon = SwfWindow("Symphony Oracle Logon")
'					OptionStep.SwfWindow("Symphony Import").Dialog("Import").WinButton("OK")
	
		If  vSymOraLogon.exist(1) Then
			fnOracleLogOn vUsername, vPassword
			fnImportProcessToDisplayReport1
			fnReportSaveSQL1(vReport1)
 @@ hightlight id_;_2294406_;_script infofile_;_ZIP::ssf126.xml_;_
			Else
				fnImportProcessToDisplayReport1
				fnReportSaveSQL1(vReport1)

		End if
		''''************************** 
End Function
''''*****************************
Function fnBluelightSettingToSQL(vBL_App_Setting_Path)
		   SystemUtil.Run vBL_App_Setting_Path
		   ''''*********************************

		Wait(10)
		Set vSymOraLogon = SwfWindow("Symphony Oracle Logon")
	With	SwfWindow("Symphony Settings")
		If  vSymOraLogon.exist(1) Then
			fnOracleLogOn vUsername, vPassword
			.SwfObject("C:\SSF\BLUELIGHT_NAG.SSF").Click 714,14
			With	.Dialog("Select a settings file")
				.WinEdit("File name:").Set vSSFConfigFileSQL
				.WinButton("Open").Click @@ hightlight id_;_2688486_;_script infofile_;_ZIP::ssf110.xml_;_
			End With
			.SwfObject("Test Connection").Click 56,11 @@ hightlight id_;_1574314_;_script infofile_;_ZIP::ssf112.xml_;_
			With	.Dialog("Symphony Settings")
				.WinButton("OK").Click @@ hightlight id_;_2230298_;_script infofile_;_ZIP::ssf113.xml_;_
				SwfWindow("Symphony Settings").SwfObject("OK").Click 32,11 @@ hightlight id_;_4785032_;_script infofile_;_ZIP::ssf114.xml_;_
				.WinButton("Yes").Click @@ hightlight id_;_27068276_;_script infofile_;_ZIP::ssf115.xml_;_
			End With
			Else
				OptionalStep.SwfWindow("Symphony Settings").SwfObject("OK").Click 32,11
		End if
	End With
End Function
'*************
Function fnBluelightSettingToORACLE(vBL_App_Setting_Path)
					SystemUtil.Run vBL_App_Setting_Path		'''*********************************       
					Wait(10)  
					If  SwfWindow("Symphony Oracle Logon").exist(1) Then
						fnOracleLogOn vUsername, vPassword
						SwfWindow("Symphony Settings").SwfObject("OK").Click 32,11
				   Else
			With	SwfWindow("Symphony Settings")
				.SwfObject("C:\SSF\BLUELIGHT_NAG.SSF").Click 714,14
				With	.Dialog("Select a settings file")
					.WinEdit("File name:").Set vSSFConfigFileORA
					.WinButton("Open").Click
				End With
				fnSymOraLogonWithinSetting vUsername, vPassword
				.SwfObject("Test Connection").Click 56,11 @@ hightlight id_;_1574314_;_script infofile_;_ZIP::ssf112.xml_;_
				With	.Dialog("Symphony Settings")
					.WinButton("OK").Click @@ hightlight id_;_2230298_;_script infofile_;_ZIP::ssf113.xml_;_
					SwfWindow("Symphony Settings").SwfObject("OK").Click 32,11 @@ hightlight id_;_4785032_;_script infofile_;_ZIP::ssf114.xml_;_
					.WinButton("Yes").Click @@ hightlight id_;_27068276_;_script infofile_;_ZIP::ssf115.xml_;_
				End With
			End With
	End if
		   
End Function

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

Function fnImportProcessToDisplayReport1()
'				SwfWindow("Symphony Import").Activate
				' In case Default screen appears with 'NEXT' button after fresh installation
'				OptionalStep.SwfWindow("Symphony Import").SwfObject("Next >").Click 36,11
			If Trim(OptionalStep.SwfWindow("Symphony Import").SwfObject("Welcome to the Symphony").GetROProperty("Text")) = vWelcomePage Then
				
				SwfWindow("Symphony Import").SwfObject("Next >").Click 36,11
				print "Welcome page found in COU"
			End If
				SwfWindow("Symphony Import").SwfObject("OS Insight - AddressBase").Click 51,33
				Wait(3)
'				********** If IgnorFailed Data and Reprocess screen appear ***
				if (instr(SwfWindow("Symphony Import").SwfObject("IgnoreFailedData").GetROProperty("Text"), "last") <> 0) then 
				
						wait(3)
						SwfWindow("Symphony Import").SwfObject("grpProcess").Click 272,38 @@ hightlight id_;_1114848_;_script infofile_;_ZIP::ssf130.xml_;_
						wait(2)
						SwfWindow("Symphony Import").SwfObject("Next >").Click 26,12 @@ hightlight id_;_10485850_;_script infofile_;_ZIP::ssf131.xml_;_
						print "Got ignore"
				end if	
				
'				*******
				
				SwfWindow("Symphony Import").SwfObject("Next >").Click 36,11
				SwfWindow("Symphony Import").SwfObject("Next >").Click 36,11
				SwfWindow("Symphony Import").SwfObject("Next >").Click 36,11
				SwfWindow("Symphony Import").SwfObject("Next >").Click 36,11
				
				OptionalStep.SwfWindow("Symphony Import").Dialog("Import").WinButton("OK").click
				OptionalStep.Window("Symphony Import").Dialog("Import").WinButton("OK").Click
'				*************** Logic when csv file import completed then fnWaitTilExist will be true ****
				If  fnWaitTilExists = True then
					SwfWindow("Symphony Import").SwfObject("Display Report").Click 45,11
				end if
'				SwfWindow("Symphony Import").Close
End Function

Function fnReportSave(vReport) 
'			vReport = "C:\Automation\Bluelight iExchange\Local_CSVFolder\T002\Report"
			Set DeviceReplay = CreateObject("Mercury.DeviceReplay")
			Set WshShell1 = CreateObject("Wscript.Shell")
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
			Set WshShell1 = CreateObject("Wscript.Shell")
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


'systemutil.Run "C:\Program Files\Aligned Assets Limited\Symphony Bluelight Gazetteer\Import.exe"
wait(5)
SwfWindow("Symphony Import").SwfObject("Next >").Click 50,15 @@ hightlight id_;_10485850_;_script infofile_;_ZIP::ssf127.xml_;_
SwfWindow("Symphony Import").SwfObject("OS Insight - AddressBase").Click 175,35 @@ hightlight id_;_1509094_;_script infofile_;_ZIP::ssf128.xml_;_
SwfWindow("Symphony Import").SwfObject("Next >").Click 34,11 @@ hightlight id_;_10485850_;_script infofile_;_ZIP::ssf129.xml_;_
if (instr(SwfWindow("Symphony Import").SwfObject("IgnoreFailedData").GetROProperty("Text"), "last") <> 0) then 
'swfWindow("Symphony Import").SwfObject("The last time you ran_2").GetROProperty("Text")
		wait(3)
		SwfWindow("Symphony Import").SwfObject("grpProcess").Click 272,38 @@ hightlight id_;_1114848_;_script infofile_;_ZIP::ssf130.xml_;_
		wait(2)
		SwfWindow("Symphony Import").SwfObject("Next >").Click 26,12 @@ hightlight id_;_10485850_;_script infofile_;_ZIP::ssf131.xml_;_
		print "Got ignore"
end if


