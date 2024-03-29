Dim App 'As Application
Set App = CreateObject("QuickTest.Application")
App.Launch
App.Visible = True
App.Test.Settings.Launchers("Web").Active = False
App.Test.Settings.Launchers("Web").Browser = "IE"
App.Test.Settings.Launchers("Web").Address = "http://newtours.demoaut.com "
App.Test.Settings.Launchers("Web").CloseOnExit = True
App.Test.Settings.Launchers("Windows Applications").Active = True
App.Test.Settings.Launchers("Windows Applications").Applications.RemoveAll
App.Test.Settings.Launchers("Windows Applications").RecordOnQTDescendants = True
App.Test.Settings.Launchers("Windows Applications").RecordOnExplorerDescendants = False
App.Test.Settings.Launchers("Windows Applications").RecordOnSpecifiedApplications = True
App.Test.Settings.Run.IterationMode = "rngAll"
App.Test.Settings.Run.StartIteration = 1
App.Test.Settings.Run.EndIteration = 1
App.Test.Settings.Run.ObjectSyncTimeOut = 20000
App.Test.Settings.Run.DisableSmartIdentification = False
App.Test.Settings.Run.OnError = "Dialog"
App.Test.Settings.Resources.DataTablePath = "<Default>"
App.Test.Settings.Resources.Libraries.RemoveAll
App.Test.Settings.Resources.Libraries.Add("Globals.qfl")
App.Test.Settings.Resources.Libraries.Add("Log.qfl")
App.Test.Settings.Resources.Libraries.Add("ObjActions.qfl")
App.Test.Settings.Resources.Libraries.Add("ObjUtil.qfl")
App.Test.Settings.Resources.Libraries.Add("Verification.qfl")
App.Test.Settings.Resources.Libraries.Add("Utility.qfl")
App.Test.Settings.Resources.Libraries.Add("Synchronize.qfl")
App.Test.Settings.Resources.Libraries.Add("ValidationObjects.qfl")
App.Test.Settings.Resources.Libraries.Add("ConfigParams.qfl")
App.Test.Settings.Resources.Libraries.Add("RuntimeParams.qfl")
App.Test.Settings.Resources.Libraries.Add("Variables.qfl")
App.Test.Settings.Resources.Libraries.Add("Init.qfl")
App.Test.Settings.Resources.Libraries.Add("SMTP.qfl")
App.Test.Settings.Resources.Libraries.Add("UserPermissions.qfl")
App.Test.Settings.Resources.Libraries.Add("ApptActivity.qfl")
App.Test.Settings.Resources.Libraries.Add("Login_Logout.qfl")
App.Test.Settings.Resources.Libraries.Add("Workflow.qfl")
App.Test.Settings.Resources.Libraries.Add("BBAObjUtil.qfl")
App.Test.Settings.Web.BrowserNavigationTimeout = 60000
App.Test.Settings.Web.ActiveScreenAccess.UserName = ""
App.Test.Settings.Web.ActiveScreenAccess.Password = ""
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' System Local Monitoring settings
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
App.Test.Settings.LocalSystemMonitor.Enable = false
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Log Tracking settings
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
With App.Test.Settings.LogTracking 
	.IncludeInResults = False 
	.Port = 18081 
	.IP = "127.0.0.1" 
	.MinTriggerLevel = "ERROR" 
	.EnableAutoConfig = False 
	.RecoverConfigAfterRun = False 
	.ConfigFile = "" 
	.MinConfigLevel = "WARN" 
End With
