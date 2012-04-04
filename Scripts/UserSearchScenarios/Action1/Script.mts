'Test Data
'DataTable.ImportSheet "C:\automation\Data\UserSearchScenarios.xls",1,Global

Init()
schedulerLogin tdGetUserName,tdGetPassword
goToLoginView tdGetView
goToModule tdGetModule
selectBranch  "", tdGetBranchName,tdGetLocationType

searchInUserManager tdGetTabOption
verifySearchResults tdGetTabOption,tdGetSearchDataDict
