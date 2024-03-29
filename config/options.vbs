Dim App 'As Application
Set App = CreateObject("QuickTest.Application")
App.Launch
App.Visible = True
App.Options.DisableVORecognition = False
App.Options.AutoGenerateWith = False
App.Options.WithGenerationLevel = 2
App.Options.TimeToActivateWinAfterPoint = 500
App.Options.SaveLoadAndMonitorData = False
App.Options.Run.RunMode = "Normal"
App.Options.Run.ViewResults = False
App.Options.Run.StepExecutionDelay = 0
App.Options.Run.MovieCaptureForTestResults = "Never"
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
App.Options.Web.PageCreationMode = "Description"
App.Options.Web.CreatePageUsingUserData = "Get Post"
App.Options.Web.CreatePageUsingNonUserData = ""
App.Options.Web.CreatePageUsingAdditionalInfo = True
App.Options.Web.FrameCreationMode = "Description"
App.Options.Web.CreateFrameUsingUserData = "Get Post"
App.Options.Web.CreateFrameUsingNonUserData = ""
App.Options.Web.CreateFrameUsingAdditionalInfo = True
App.Options.Web.UseAutoXPathIdentifiers = True
App.Options.WindowsApps.AttachedTextRadius = 35
App.Options.WindowsApps.AttachedTextArea = "TopLeft"
App.Options.WindowsApps.ExpandMenuToRetrieveProperties = True
App.Options.WindowsApps.NonUniqueListItemRecordMode = "ByName"
App.Options.WindowsApps.RecordOwnerDrawnButtonAs = "PushButtons"
App.Options.WindowsApps.ForceEnumChildWindows = 0
App.Options.WindowsApps.ClickEditBeforeSetText = 0
App.Options.WindowsApps.VerifyMenuInitEvent = 0
App.Options.TextRecognitionOrder = "APIThenOCR"
App.Options.TextRecognitionBlockType = "Multiple"
App.Options.TextRecognitionLanguages = "English"
App.Options.DisplayKeywordView = True
App.Options.AutoParameterizeSteps = False
App.Options.AutoParameterType = "Data Table"
App.Folders.RemoveAll
App.Folders.Add("C:\Automation\Data")
App.Folders.Add("C:\Automation\Lib\BBA")
App.Folders.Add("C:\Automation\Lib\Common")
App.Folders.Add("C:\Automation\Lib\Framework")
App.Folders.Add("C:\Automation\Repositories")
