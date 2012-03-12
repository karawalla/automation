'Test Data
'DataTable.ImportSheet "C:\Chandra\Scheduler_Automation\DataSheet.xls",1,Global
DataTable.ImportSheet "C:\SmokeScripts\ObjectValidations.xls",1,Global
Init()

'schedulerLogin
loginData = Split(DataTable.Value("LoginData",Global),":")
userName = loginData(0)
password = loginData(1)

'goToLoginView
viewName = DataTable.Value("View",Global)

'selectBranch
branchName = DataTable.Value("BranchName",Global)

locationtype = DataTable.Value("LocationType",Global)
addToCache "DT_Location",locationtype

objectToValidate = DataTable.Value("ObjToValidate",Global)
'*************************************************************
schedulerLogin userName,password
goToLoginView viewName
goToModule moduleName
selectBranch  branchType, branchName

Select Case objectToValidate
	Case "WalkIn"
			checkIsNotExist Browser("Smart Lobby").Page("Smart Lobby").Frame("Frame_ActivityManager").WebElement("Walk-In")
	Case "NoShow"
			checkIsNotExist Browser("Smart Lobby").Page("Smart Lobby").Frame("Frame_ActivityManager").WebElement("No Show")
End Select
endTest
'##############################################################
