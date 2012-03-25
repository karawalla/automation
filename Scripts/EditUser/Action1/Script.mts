'Test Data
DataTable.ImportSheet "C:\automation\Data\UserPermissions.xls","EditUser",Global
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

'Actual user Data to get Record
userFirstName = DataTable.Value("ActualUserFirstName",Global)
userLastName = DataTable.Value("ActualUserLastName",Global)

'Set Data

setUserFirstName = DataTable.Value("SetUserFirstName",Global)
setUserLastName = DataTable.Value("SetUserLastName",Global)
setUserEmailID = DataTable.Value("SetUserEmailID",Global)
setCaid =DataTable.Value("SetCaid",Global)
setUserRole = DataTable.Value("SetUserRole",Global)
setUserGroup = DataTable.Value("SetUserGroup",Global)

addToCache "DT_UserFirstName", DataTable.Value("ActualUserFirstName",Global)
addToCache "DT_UserLastName", DataTable.Value("ActualUserLastName",Global)

addToUserInfo "DT_UserFirstName",setUserFirstName
addToUserInfo "DT_UserLastName", setUserLastName
addToUserInfo "DT_UserEmail", setUserEmailID
addToUserInfo "DT_UserCAID", setCaid
addToUserInfo "DT_UserRole" , setUserRole
addToUserInfo "DT_UserGroup", setUserGroup

tabOption = DataTable.Value("TabOption",Global)
addToUserInfo "DT_TabOption", tabOption

setCurrentView viewName
setCurrentModule moduleName
addToCache "DT_Location",locationtype

edit_UserLocations = DataTable.Value("EditUserLocations",Global)
edit_UserDetails = DataTable.Value("EditUserDetails",Global)

isErrorExpected = DataTable.Value("IsErrorExpected",Global)
'*************************************************************
schedulerLogin userName,password
goToLoginView viewName
goToModule moduleName
selectBranch  branchType, branchName

If Trim(edit_UserDetails)<>"" Then
	If Trim(UCase(edit_UserDetails)) = "TRUE" Then
		selectUserRecord userFirstName,userLastName
		goToEditUser
		addEditUserDetails
		saveUser
		'Check if Error is Expected or not
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
		Else
			checkAddEditUserSuccessful
			verifyUserRecord
		End If
	End If
End If

If Trim(edit_UserLocations)<>"" Then
	If Trim(UCase(edit_UserLocations)) = "TRUE" Then
		set_UserLocations = Array(DataTable.Value("SetUserLocation1",Global),DataTable.Value("SetUserLocation2",Global))		
		selectUserRecord userFirstName,userLastName
		goToEditUser
		editUserLocation set_UserLocations
		saveUser
		checkAddEditUserSuccessful
		For Each loc in set_UserLocations
			selectBranch branchType, loc
			goToModule moduleName
			If Trim(UCase(edit_UserDetails)) = "TRUE" Then
				selectUserRecord setUserFirstName,setUserLastName 
			Else
				selectUserRecord userFirstName,userLastName
			End If
		Next
	End If
End If
