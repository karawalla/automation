'Test Data
DataTable.ImportSheet "C:\automation\Data\UserPermissions.xls","SearchScenarios",Global
Init()

'schedulerLogin
loginData = Split(DataTable.Value("LoginData",Global),":")
userName = loginData(0)
password = loginData(1)

'goToLoginView
viewName = DataTable.Value("View",Global)

'selectBranch
branchName = DataTable.Value("BranchName",Global)

'branchType = DataTable.Value("BranchType",Global)
locationtype = DataTable.Value("LocationType",Global)

'goToModule
moduleName = DataTable.Value("Module",Global)

'setCurrentWorkFlow workFlow
userFirstName = DataTable.Value("UserFirstName",Global)
userLastName = DataTable.Value("UserLastName",Global)
userEmailID = DataTable.Value("UserEmailID",Global)
caid =DataTable.Value("Caid",Global)
tabOption = DataTable.Value("TabOption",Global)

addToUserInfo "DT_UserFirstName",userFirstName
addToUserInfo "DT_UserLastName", userLastName
addToUserInfo "DT_UserEmail", userEmailID
addToUserInfo "DT_UserCAID", caid
addToUserInfo "DT_TabOption", tabOption

setCurrentView viewName
setCurrentModule moduleName
addToCache "DT_Location",locationtype
'*************************************************************
schedulerLogin userName,password
goToLoginView viewName
goToModule moduleName
selectBranch  branchType, branchName

searchInUserManager
verifySearchResults
