'Test Data
'DataTable.ImportSheet "C:\Chandra\Scheduler_Automation\DataSheet.xls",1,Global
DataTable.ImportSheet "C:\SmokeScripts\SearchScenarios.xls",1,Global
Init()

'schedulerLogin
loginData = Split(DataTable.Value("LoginData",Global),":")
userName = loginData(0)
password = loginData(1)

'goToLoginView
viewName = DataTable.Value("View",Global)

'selectBranch
branchName = DataTable.Value("BranchName",Global)

searchCriteria = DataTable.Value("SearchCriteria",Global)
inputData = DataTable.Value("SearchDataInput",Global)

'*************************************************************
schedulerLogin userName,password
goToLoginView viewName
'goToModule moduleName
selectBranch  branchType, branchName

Select Case searchCriteria
	Case "SearchByName"
		searchByName(inputData)
	Case "SearchByCustomer"
		searchByCustomer(inputData)
	Case "SearchByRole"
		searchByRole(inputData)
	Case "SearchByAssociateName"
		searchByAssociateName(inputData)
End Select
endTest
'##############################################################