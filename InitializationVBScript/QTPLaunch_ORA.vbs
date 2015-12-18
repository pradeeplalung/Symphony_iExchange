Dim qtApp
Dim qtTest

'Create the QTP Application object
Set qtApp = CreateObject("QuickTest.Application", "PL-Win7Dual")

'If QTP is notopen then open it
If  qtApp.launched <> True then 

qtApp.Launch 

End If 

'Make the QuickTest application visible
qtApp.Visible = True

'Set QuickTest run options
'Instruct QuickTest to perform next step when error occurs

qtApp.Options.Run.ImageCaptureForTestResults = "OnError"
qtApp.Options.Run.RunMode = "Normal"
qtApp.Options.Run.StepExecutionDelay = 200
qtApp.Options.Run.ViewResults = False

'Open the test in read-only mode
qtApp.Open "C:\Automation\Sym_iEx\ServiceOperation\Symphony_ImportIX_ORA", True 
'qtApp.Open "C:\Automation\BL_iEx\ServiceOperation\\Auto_Service_Driver_Main_ORA1", True 


'set run settings for the test
Set qtTest = qtApp.Test

'Instruct QuickTest to perform next step when error occurs
qtTest.Settings.Run.OnError = "NextStep" 

'Run the test
qtTest.Run

'Check the results of the test run
'MsgBox qtTest.LastRunResults.Status

' Close the test
qtTest.Close 

'Close QTP
qtApp.quit

'Release Object
Set qtTest = Nothing
Set qtApp = Nothing 



