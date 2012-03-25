'Test Data
DataTable.ImportSheet "C:\automation\Data\UserPermissions.xls",1,Global
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
userRole = DataTable.Value("UserRole",Global)
userGroup = DataTable.Value("UserGroup",Global)
tabOption = DataTable.Value("TabOption",Global)

addToUserInfo "DT_UserFirstName",userFirstName
addToUserInfo "DT_UserLastName", userLastName
addToUserInfo "DT_UserEmail", userEmailID
addToUserInfo "DT_UserCAID", caid
addToUserInfo "DT_UserRole" , userRole
addToUserInfo "DT_UserGroup", userGroup
addToUserInfo "DT_TabOption", tabOption

isErrorExpected = DataTable.Value("IsErrorExpected",Global)

setCurrentView viewName
setCurrentModule moduleName
addToCache "DT_Location",locationtype
'*************************************************************
schedulerLogin userName,password
goToLoginView viewName
goToModule moduleName
selectBranch  branchType, branchName

'Add User WorkFlow
checkAndAddUser

'Check if Error is Expected or not
If retrieveFromCache("DT_UserRecordFound")<> True Then
	If trim(UCase(isErrorExpected))="TRUE" Then
		isErrorExists = checkCAIDAlreadyExists()
		isDialogExists = checkErrorDailogExists()
		If isErrorExists Then
			logPass "Error Msg :Add Edir Uer - CAID already exists.Please enter another CAID"
        ElseIf isDialogExists Then
			errText = retrieveFromCache("DT_DailogText")
			logPass "Error Msg: "& errText
		ElseIf Browser("AddEditUser").Page("AddEditUser").Exist(1) Then
			logFatal "Unknown Error - AditEdit user still exists and not Closed."
		Else
			logFail "No Error Found.Please check the user input Data."
		End If
	'Else Normal WorkFlow
	Else
		addToCache "DT_UserRecordFound", False
		checkAddEditUserSuccessful
		verifyUserRecord
	End If
End If
