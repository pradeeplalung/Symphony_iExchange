'Test Object : Import COU csv file in Oracle and SQL environment

'Option explicit
Dim vCSVPath,vCSVFile,vSSFConfigFile,vSSFConfigFileORA,vSSFConfigFileSQL, vSearchImportFolderLine,vReplaceImportFolderLine,vSymOraLogon, vPattern
Dim vSym_App_Path,vSym_App_Setting_Path,vTP_Path ,vTP,vTCD1,vTPS,vTSEC, vUsername,vPassword
Dim WshShell,DeviceReplay,vReport,vReport1,WshShell1

Dim vRowControl
'*********** Assigned  Environment variable  values dynamically from XML file *************
Environment.LoadFromFile("C:\Automation\Sym_iEx\Xml_File\IX_TP01_T001.xml")

vSym_App_Path = Environment.Value("vSym_App_Path")	

vSym_App_Setting_Path = Environment.Value("vSym_App_Setting_Path")
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

vSSFConfigFile =Trim(fnReadFromExcel(vTP_Path,vTPS,6,2))   '"C:\SSF\TRAINING_1.SSF"
vSSFConfigFile =Split(vSSFConfigFile,";")
vSSFConfigFileORA = Trim(vSSFConfigFile(0))
vSSFConfigFileSQL = Trim(vSSFConfigFile(1))
print "vSSFConfigFileORA  :" &vSSFConfigFileORA
print "vSSFConfigFileSQL  :" &vSSFConfigFileSQL
vPattern = ".*Import Folder.*"
vCSVPath = Trim(fnReadFromExcel(vTP_Path,vTCD1,2,2))
print "vCSVPath  :"&vCSVPath ' Bulk import
vCSVFile = Trim(fnReadFromExcel(vTP_Path,vTCD1,6,2))
print "vCSVFile  :"&vCSVFile ' Non-bulk import


vReplaceImportFolderLine = "<entry name="&chr(34)&"Import Folder"&chr(34)&">"&vCSVPath&"</entry>"
print "vReplaceImportFolderLine  :"&vReplaceImportFolderLine

vRowControl =  fnReadFromExcel(vTP_Path,vTCD1,3,4)
print "vRowControl  :"&vRowControl

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
		        If  fnReadFromExcel(vTP_Path,vTPS,21,2) = "Yes" and fnReadFromExcel(vTP_Path,vTPS,22,2) = "No" and fnReadFromExcel(vTP_Path,vTCD1,3,3) = "Yes" Then
											On error resume next
											fnT001 vRowControl,vReplaceImportFolderLine
				End If
End If
'''****************** For SQL Environment************************
dim x,y,z
x = fnReadFromExcel(vTP_Path,vTPS,21,2)
print "x :"&x
y = fnReadFromExcel(vTP_Path,vTPS,22,2)
print "y :"&y
z = fnReadFromExcel(vTP_Path,vTCD1,4,3)
print "z :"&z

If  fnReadFromExcel(vTP_Path,vTPS,21,2) = "No" and fnReadFromExcel(vTP_Path,vTPS,22,2) = "Yes" and  fnReadFromExcel(vTP_Path,vTCD1,4,3) = "Yes" Then
											print "Entry Logic "
											On error resume next
											fnT002 vRowControl+1,vReplaceImportFolderLine
End If

'***************
Function fnT001(vRowControl,vReplaceImportFolderLine)
					 fnSym_iExchangeSettingToORACLE(vSym_App_Setting_Path)
					vReport1 =Trim(fnReadFromExcel(vTP_Path,vTCD1,vRowControl,2)) ' "C:\Automation\BL_iEx\CSV\Import Files\COU_Report\ORACLE"
					'msgbox vReport1
					fnDeleteFileInFolder vReport1&"\"
					vSearchImportFolderLine =  fnSearchLineOfPatternInFile(vSSFConfigFileORA,vPattern)
					fnReplaceLineWithSpecificValue vSSFConfigFileORA,vSearchImportFolderLine,vReplaceImportFolderLine,vSSFConfigFileORA
					
					'*********************************
					 fnStartIBLImport vSym_App_Path
					''''*********************************
					Wait(10)
					Set vSymOraLogon = SwfWindow("Symphony Oracle Logon")
					If  vSymOraLogon.exist(1) Then
								fnOracleLogOn vUsername, vPassword
'								 fnImportProcessToDisplayReport1
								fnImportProcessToDisplayReport1_WithoutBulk(vCSVFile)
								fnReportSaveSQL1(vReport1)

					End if
'************************** 
End Function
Function fnT002(vRowControl,vReplaceImportFolderLine)
					print "in no bulk import module"
					print "vRowControl  :"&vRowControl
					fnSym_iExchangeSettingToSQL(vSym_App_Setting_Path)
					vReport1 =Trim(fnReadFromExcel(vTP_Path,vTCD1,vRowControl,2)) '  "C:\Automation\BL_iEx\CSV\Import Files\COU_Report\SQL"
					fnDeleteFileInFolder vReport1&"\"
					vSearchImportFolderLine =  fnSearchLineOfPatternInFile(vSSFConfigFileSQL,vPattern)
					print "vSearchImportFolderLine  :"&vSearchImportFolderLine
					fnReplaceLineWithSpecificValue vSSFConfigFileSQL,vSearchImportFolderLine,vReplaceImportFolderLine,vSSFConfigFileSQL

					
''''					'*********************************
					 fnStartIBLImport vSym_App_Path
''''					''''*********************************
					Wait(10)
					Set vSymOraLogon = SwfWindow("Symphony Oracle Logon")
'					OptionStep.SwfWindow("Symphony Import").Dialog("Import").WinButton("OK")
	
		If  vSymOraLogon.exist(1) Then
			fnOracleLogOn vUsername, vPassword
			fnImportProcessToDisplayReport1_WithoutBulk(vCSVFile)
			fnReportSaveSQL1(vReport1)
 @@ hightlight id_;_2294406_;_script infofile_;_ZIP::ssf126.xml_;_
			Else
				fnImportProcessToDisplayReport1_WithoutBulk(vCSVFile)
				fnReportSaveSQL1(vReport1)

		End if
		''''************************** 
End Function
''''*****************************
Function fnSym_iExchangeSettingToSQL(vSym_App_Setting_Path)
		   SystemUtil.Run vSym_App_Setting_Path
		   ''''*********************************

		Wait(10)
		Set vSymOraLogon = SwfWindow("Symphony Oracle Logon")
	With	SwfWindow("Symphony Settings")
		If  vSymOraLogon.exist(1) Then
			fnOracleLogOn vUsername, vPassword
			.SwfObject("C:\SSF\TRAINING_1.SSF").Click 714,14
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
Function fnSym_iExchangeSettingToORACLE(vSym_App_Setting_Path)
					SystemUtil.Run vSym_App_Setting_Path		'''*********************************       
					Wait(10)  
					If  SwfWindow("Symphony Oracle Logon").exist(1) Then
						fnOracleLogOn vUsername, vPassword
						SwfWindow("Symphony Settings").SwfObject("OK").Click 32,11
				   Else
			With	SwfWindow("Symphony Settings")
				.SwfObject("C:\SSF\TRAINING_1.SSF").Click 714,14
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


'**********************************


'systemutil.Run "C:\Program Files\Aligned Assets Limited\Symphony Gazetteer\Import.exe"
'wait(5)
''
'vWelcomePage = "Welcome to the Symphony Gazetteer Import Wizard"
'vCSVFile = "C:\Automation\SinglePoint\LLPG\WebSite\CSVFile\6815_20091016_01.csv"
'fnImportProcessToDisplayReport1_WithoutBulk vCSVFile

Function fnImportProcessToDisplayReport1_WithoutBulk(vCSVFile)
	
		If Trim(OptionalStep.SwfWindow("Symphony Import").SwfObject("Welcome to the Symphony").GetROProperty("Text")) = vWelcomePage Then
				
'				SwfWindow("Symphony Import").SwfObject("Next >").Click 36,11
				
				SwfWindow("Symphony Import").SwfObject("Next >").Click 31,12 @@ hightlight id_;_198558_;_script infofile_;_ZIP::ssf132.xml_;_
				print "Welcome page found in COU"
			End If				
			
				SwfWindow("Symphony Import").SwfObject("Partial_NLPG").Click 44,13
				SwfWindow("Symphony Import").SwfObject("Next >").Click 31,12 @@ hightlight id_;_198558_;_script infofile_;_ZIP::ssf134.xml_;_
				Wait(3)
				
				
'				********** If IgnorFailed Data and Reprocess screen appear ***
				if (instr(OptionalStep.SwfWindow("Symphony Import").SwfObject("IgnoreFailedData_NLPG").GetROProperty("Text"), "last") <> 0) then 
				
						wait(3)
						SwfWindow("Symphony Import").SwfObject("grpProcess").Click 196,33 @@ hightlight id_;_1114848_;_script infofile_;_ZIP::ssf130.xml_;_
						wait(2)
						SwfWindow("Symphony Import").SwfObject("Next >").Click 26,12 @@ hightlight id_;_10485850_;_script infofile_;_ZIP::ssf131.xml_;_
						print "Got ignore"
				end if	
				wait(2)
				SwfWindow("Symphony Import").SwfObject("chkFlagAsImported").Click 83,12
				SwfWindow("Symphony Import").SwfObject("Next >").Click 31,12
				SwfWindow("Symphony Import").SwfObject("Next >").Click 31,12
				SwfWindow("Symphony Import").SwfObject("Next >").Click 31,12
'
''''**************************** Selecting the Non-Bulk cvs file 
				SwfWindow("Symphony Import").SwfObject("txtFilename").Click 383,13
				wait(2)
				SwfWindow("Symphony Import").Dialog("Select NLPG CSV Import").WinEdit("File name:").Set vCSVFile
				SwfWindow("Symphony Import").Dialog("Select NLPG CSV Import").WinButton("Open").Click
'				
'				****************************
				SwfWindow("Symphony Import").SwfObject("Next >").Click 31,12
				wait(2)
				SwfWindow("Symphony Import").SwfObject("Next >").Click 31,12
				
				OptionalStep.SwfWindow("Symphony Import").Dialog("Import").WinButton("OK").click
				OptionalStep.Window("Symphony Import").Dialog("Import").WinButton("OK").Click
'				*************** Logic when csv file import completed then fnWaitTilExist will be true ****
				If  fnWaitTilExists = True then
'					SwfWindow("Symphony Import").SwfObject("Display Report").Click 45,11#
				SwfWindow("Symphony Import").SwfObject("Display Full Report").Click 48,11 @@ hightlight id_;_788204_;_script infofile_;_ZIP::ssf151.xml_;_
'				SwfWindow("Symphony Import").SwfWindow("Import Report").SwfObject("Save Report").Click 49,10 @@ hightlight id_;_788176_;_script infofile_;_ZIP::ssf152.xml_;_
'				SwfWindow("Symphony Import").SwfWindow("Import Report").Dialog("Save Import Report As").WinButton("Save").Click @@ hightlight id_;_1704792_;_script infofile_;_ZIP::ssf153.xml_;_
'				SwfWindow("Symphony Import").SwfWindow("Import Report").Dialog("Save Report").WinButton("OK").Click
							
				end if	
End Function	
'				
'vReport = "C:\Automation\Sym_iEx\CSV\Import Files\NLPG_Report\SQL"
'fnReportSaveSQL1 vReport

Function fnReportSaveSQL1(vReport) 

			Set DeviceReplay = CreateObject("Mercury.DeviceReplay")
			Set WshShell1 = CreateObject("Wscript.Shell")
	
'			******** Saving excel Report *****
			SwfWindow("Symphony Import").SwfWindow("Import Report").SwfObject("Export Report").Click 58,11
			DeviceReplay.SendString vReport&"\Import Errors.xls" 
			Wait(3)
			SwfWindow("Symphony Import").SwfWindow("Import Report").Dialog("Export Import Report As").WinButton("Save").Click
'			*********** Saving pdf report ******
			SwfWindow("Symphony Import").SwfWindow("Import Report").SwfObject("Save Report").Click 49,10 @@ hightlight id_;_788176_;_script infofile_;_ZIP::ssf152.xml_;_
			SwfWindow("Symphony Import").SwfWindow("Import Report").Dialog("Save Import Report As").WinButton("Save").Click @@ hightlight id_;_1704792_;_script infofile_;_ZIP::ssf153.xml_;_
			SwfWindow("Symphony Import").SwfWindow("Import Report").Dialog("Save Report").WinButton("OK").Click

'			**********
			SwfWindow("Symphony Import").SwfWindow("Import Report").SwfObject("Exit").Click 41,12
			Set DeviceReplay = Nothing
			Set WshShell1 = Nothing
End Function


 @@ hightlight id_;_11338648_;_script infofile_;_ZIP::ssf155.xml_;_