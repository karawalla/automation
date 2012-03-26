Init()

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
