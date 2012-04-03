
Init()

'Test Data
'DataTable.ImportSheet "C:\automation\Data\AddUser.xls",1,Global

'Init()
'schedulerLogin tdGetUserName,tdGetPassword
'goToLoginView tdGetView
'goToModule tdGetModule
'selectBranch  "", tdGetBranchName

schedulerLogin "masteradmin@ncr.com","masteradmin"
goToLoginView "Lobby"
goToModule "usermanager"
selectBranch  "", "Arboretum"

'Add User WorkFlow
'checkAndAddUser
goToAddUser(tdGetTabOption)
addEditUserDetails tdGetTabOption,tdGetAddUserInfoDict
saveUser

'Check if Error is Expected or not
If trim(UCase(tdGetAddUserInfoDict.Item("IsError_Expected")))="TRUE" Then
	isErrorExists = checkCAIDAlreadyExists()
	isDialogExists = checkErrorDailogExists()
	If isErrorExists Then
		logPass "Error Msg :Add Edir Uer - CAID already exists.Please enter another CAID"
	ElseIf isDialogExists Then
		errText = retrieveFromCache("Dailog_Text")
		logPass "Error Msg: "& errText
	ElseIf Browser("AddEditUser").Page("AddEditUser").Exist(1) Then
		logFatal "Unknown Error - AditEdit user still exists and not Closed."
	Else
		logFail "No Error Found.Please check the user input Data."
	End If
Else
	'Else Normal WorkFlow
	checkAddEditUserSuccessful
	verifyAddUserRecord tdGetAddUserInfoDict
End If