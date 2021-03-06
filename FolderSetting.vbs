Dim App 'As Application
Set App = CreateObject("QuickTest.Application")
App.Launch
App.Visible = True
App.Options.DisableVORecognition = False
App.Options.AutoGenerateWith = False
App.Options.WithGenerationLevel = 2
App.Options.TimeToActivateWinAfterPoint = 500
App.Options.SaveLoadAndMonitorData = False
App.Options.TextRecognitionOrder = "OCROnly"
App.Options.TextRecognitionLanguages = "English"
App.Options.TextRecognitionBlockType = "Multiple"
App.Options.DisplayKeywordView = True
App.Options.AutoParameterizeSteps = False
App.Options.AutoParameterType = "Data Table"
App.Options.ActiveScreen.CapturedFilesStorage = "Compressed"
App.Options.ActiveScreen.CaptureLevel = "Minimum"
App.Options.ActiveScreen.Web.CaptureOriginalHTMLSource = False
App.Options.ActiveScreen.Web.ActiveScripts = "Automatic"
App.Options.ActiveScreen.Web.LoadImages = True
App.Options.ActiveScreen.Web.LoadJavaApplets = False
App.Options.ActiveScreen.Web.LoadActiveXControls = False
App.Options.ActiveScreen.Web.LoadingTimeout = 30
App.Options.Run.ImageCaptureForTestResults = "OnError"
App.Options.Run.MovieCaptureForTestResults = "Never"
App.Options.Run.MovieSegmentSize = 2048
App.Options.Run.RunMode = "Normal"
App.Options.Run.SaveMovieOfEntireRun = True
App.Options.Run.StepExecutionDelay = 200
App.Options.Run.ViewResults = False
App.Options.Run.AutoExportReportConfig.AutoExportResults = False
App.Options.Run.AutoExportReportConfig.StepDetailsReport = True
App.Options.Run.AutoExportReportConfig.DataTableReport = True
App.Options.Run.AutoExportReportConfig.LogTrackingReport = True
App.Options.Run.AutoExportReportConfig.ScreenRecorderReport = True
App.Options.Run.AutoExportReportConfig.SystemMonitorReport = True
App.Options.Run.AutoExportReportConfig.ExportLocation = ""
App.Options.Run.AutoExportReportConfig.UserDefinedXSL = ""
App.Options.Run.AutoExportReportConfig.StepDetailsReportType = "HTML"
App.Options.Run.AutoExportReportConfig.StepDetailsReportFormat = "Short"
App.Options.Run.AutoExportReportConfig.ExportForFailedRunsOnly = True
App.Options.Run.ScreenRecorder.DeactivateShowWindowContents = True
App.Options.Run.ScreenRecorder.RecordSound = False
App.Options.Run.ScreenRecorder.SetPlainWallpaper = True
App.Options.WindowsApps.AttachedTextRadius = 35
App.Options.WindowsApps.AttachedTextArea = "TopLeft"
App.Options.WindowsApps.ExpandMenuToRetrieveProperties = True
App.Options.WindowsApps.NonUniqueListItemRecordMode = "ByName"
App.Options.WindowsApps.RecordOwnerDrawnButtonAs = "PushButtons"
App.Options.WindowsApps.ForceEnumChildWindows = 0
App.Options.WindowsApps.ClickEditBeforeSetText = 0
App.Options.WindowsApps.VerifyMenuInitEvent = 1
App.Options.Web.AddToPageLoadTime = 10
App.Options.Web.RecordCoordinates = False
App.Options.Web.RecordMouseDownAndUpAsClick = False
App.Options.Web.RecordAllNavigations = False
App.Options.Web.RecordByWinMouseEvents = ""
App.Options.Web.BrowserCleanup = False
App.Options.Web.RunOnlyClick = False
App.Options.Web.RunMouseByEvents = True
App.Options.Web.RunUsingSourceIndex = True
App.Options.Web.EnableBrowserResize = True
App.Options.Web.PageCreationMode = "URL"
App.Options.Web.CreatePageUsingUserData = "Get Post"
App.Options.Web.CreatePageUsingNonUserData = ""
App.Options.Web.CreatePageUsingAdditionalInfo = True
App.Options.Web.FrameCreationMode = "URL"
App.Options.Web.CreateFrameUsingUserData = "Get Post"
App.Options.Web.CreateFrameUsingNonUserData = ""
App.Options.Web.CreateFrameUsingAdditionalInfo = True
App.Options.Web.UseAutoXPathIdentifiers = True
App.Folders.RemoveAll
App.Folders.Add("C:\Automation\Sym_iEx\Function Library")
App.Folders.Add("C:\Automation\SinglePoint\LLPG\WebSite")
App.Folders.Add("C:\Automation\SinglePoint\SettingEnvironment")
App.Folders.Add("C:\Automation\Recovery Scenario")
App.Folders.Add("C:\Automation\SinglePoint\Recovery Scenario")
App.Folders.Add("C:\Automation\Sym_iEx\Recovery Scenario")
App.Folders.Add("C:\Automation\Sym_iEx\Driver Script")
