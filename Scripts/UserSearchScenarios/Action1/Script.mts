'Test Data
DataTable.ImportSheet "C:\automation\Data\UserSearchScenarios.xls",1,Global
Init()

schedulerLogin userName,password
goToLoginView viewName
goToModule moduleName
selectBranch  branchType, branchName

searchInUserManager tdGetTabOption
verifySearchResults tdGetTabOption,tdGetSearchDataDict
