'Test Data
'DataTable.ImportSheet "C:\automation\Data\EditUser.xls",1,Global

Init()
schedulerLogin tdGetUserName,tdGetPassword
goToLoginView tdGetView
goToModule tdGetModule
selectBranch  "", tdGetBranchName,tdGetLocationType

'Add Actual user First & Last Names to Runtime cache
addToCache "User_First_Name", tdGetActualUserFirstName
addToCache "User_Last_Name", tdGetActualUserLastName

If Trim(tdGetEditUserDetails)<>"" Then
	If Trim(UCase(tdGetEditUserDetails)) = "TRUE" Then
		selectUserRecord tdGetTabOption,tdGetActualUserFirstName,tdGetActualUserLastName		
		goToEditUser
		addEditUserDetails tdGetTabOption,tdGetEditUserInfoDict
		saveUser
	End If
End If

If Trim(tdGetEditUserLocations)<>"" Then
	If Trim(UCase(tdGetEditUserLocations)) = "TRUE" Then
		'set_UserLocations = Array(DataTable.Value("SetUserLocation1",Global),DataTable.Value("SetUserLocation2",Global))
		set_UserLocations = Array(tdGetEditUserInfoDict.Item("User_Location1"), tdGetEditUserInfoDict.Item("User_Location2"))
		selectUserRecord tdGetTabOption,tdGetActualUserFirstName,tdGetActualUserLastName
		goToEditUser
		editUserLocation set_UserLocations
		saveUser
		checkAddEditUserSuccessful
		For Each loc in set_UserLocations
			selectBranch  "", loc,tdGetLocationType			
			goToModule moduleName
			If Trim(UCase(edit_UserDetails)) = "TRUE" Then
				selectUserRecord setUserFirstName,setUserLastName 
			Else
				selectUserRecord userFirstName,userLastName
			End If
		Next
	End If
End If

'Check if Error is Expected or not
If trim(UCase(tdGetEditUserInfoDict.Item("IsError_Expected")))="TRUE" Then
	isErrorExists = checkCAIDAlreadyExists()
	isDialogExists = checkErrorDailogExists()
	If isErrorExists Then				
		logPass "Error Msg :Add Edir User - CAID already exists.Please enter another CAID"
	ElseIf isDialogExists Then
		errText = retrieveFromCache("DT_DailogText")
		logPass "Error Msg: "& errText
	ElseIf Browser("AddEditUser").Page("AddEditUser").Exist(1) Then
		logFatal "Unknown Error - AditEdit user still exists and not Closed."
	Else
		logFail "No Error Found.Please check the user input Data."
	End If
ElseIf trim(UCase(tdGetEditUserInfoDict.Item("IsError_Expected")))<>"TRUE" And Trim(UCase(tdGetEditUserLocations)) <> "TRUE" Then
	checkAddEditUserSuccessful
	verifyEditUserRecord tdGetEditUserInfoDict
End If

