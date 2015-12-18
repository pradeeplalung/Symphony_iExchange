'Option explicit
Dim vCSVPath,vSSFConfigFile,vSSFConfigFileORA,vSSFConfigFileSQL, vSearchImportFolderLine,vReplaceImportFolderLine,vSymOraLogon, vPattern
Dim vBL_App_Path,vBL_App_Setting_Path,vTP_Path ,vTP,vTCD1,vTPS,vTSEC, vUsername,vPassword
Dim WshShell,DeviceReplay,vReport,vReport1,WshShell1

Dim vRowControl
'*********** Assigned  Environment variable  values dynamically from XML file *************
Environment.LoadFromFile("C:\Automation\BL_iEx\Xml_File1\BX_TP01_T001.xml")

vBL_App_Path = Environment.Value("vBL_App_Path")	

vBL_App_Setting_Path = Environment.Value("vBL_App_Setting_Path")
vTP_Path = Environment.Value("vTP_Path")					
	
vTP = Environment.Value("vTP") ' Test Plan  Work Sheet ,
    					
vTCD1 =  Environment.Value("vTCD1")

							
vTPS = Environment.Value("vTPS")


vTSEC = Environment.Value("vTSEC")


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

vReplaceImportFolderLine = "<entry name="&chr(34)&"Import Folder"&chr(34)&">"&vCSVPath&"</entry>"

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
					fnReplaceLineWithSpecificValue vSSFConfigFileSQL,vSearchImportFolderLine,vReplaceImportFolderLine,vSSFConfigFileSQL
					
''''					'*********************************
					 fnStartIBLImport vBL_App_Path
''''					''''*********************************
					Wait(10)
					Set vSymOraLogon = SwfWindow("Symphony Oracle Logon")
					If  vSymOraLogon.exist(1) Then
								fnOracleLogOn vUsername, vPassword
								 fnImportProcessToDisplayReport1
								fnReportSaveSQL1(vReport1)
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
		If  vSymOraLogon.exist(1) Then
			fnOracleLogOn vUsername, vPassword
			SwfWindow("Symphony Settings").SwfObject("C:\SSF\BLUELIGHT_NAG.SSF").Click 714,14
			SwfWindow("Symphony Settings").Dialog("Select a settings file").WinEdit("File name:").Set vSSFConfigFileSQL
			SwfWindow("Symphony Settings").Dialog("Select a settings file").WinButton("Open").Click @@ hightlight id_;_2688486_;_script infofile_;_ZIP::ssf110.xml_;_
			SwfWindow("Symphony Settings").SwfObject("Test Connection").Click 56,11 @@ hightlight id_;_1574314_;_script infofile_;_ZIP::ssf112.xml_;_
			SwfWindow("Symphony Settings").Dialog("Symphony Settings").WinButton("OK").Click @@ hightlight id_;_2230298_;_script infofile_;_ZIP::ssf113.xml_;_
			SwfWindow("Symphony Settings").SwfObject("OK").Click 32,11 @@ hightlight id_;_4785032_;_script infofile_;_ZIP::ssf114.xml_;_
			SwfWindow("Symphony Settings").Dialog("Symphony Settings").WinButton("Yes").Click @@ hightlight id_;_27068276_;_script infofile_;_ZIP::ssf115.xml_;_
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
						SwfWindow("Symphony Settings").Dialog("Select a settings file").WinEdit("File name:").Set vSSFConfigFileORA
						SwfWindow("Symphony Settings").Dialog("Select a settings file").WinButton("Open").Click
						fnSymOraLogonWithinSetting vUsername, vPassword 		
						SwfWindow("Symphony Settings").SwfObject("Test Connection").Click 56,11 @@ hightlight id_;_1574314_;_script infofile_;_ZIP::ssf112.xml_;_
						SwfWindow("Symphony Settings").Dialog("Symphony Settings").WinButton("OK").Click @@ hightlight id_;_2230298_;_script infofile_;_ZIP::ssf113.xml_;_
						SwfWindow("Symphony Settings").SwfObject("OK").Click 32,11 @@ hightlight id_;_4785032_;_script infofile_;_ZIP::ssf114.xml_;_
						SwfWindow("Symphony Settings").Dialog("Symphony Settings").WinButton("Yes").Click @@ hightlight id_;_27068276_;_script infofile_;_ZIP::ssf115.xml_;_
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