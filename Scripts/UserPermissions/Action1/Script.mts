'Test Data
DataTable.ImportSheet "C:\SmokeScripts\UserPermissions.xls",1,Global
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

'set workFlow
'workFlow = DataTable.Value("WorkFlow", Global)
'setCurrentWorkFlow workFlow
userFirstName = "Chandra"
userLastName = "Mouli"
userEmailID = "mouliayyala@gmail.com"
caid = "Test@123"
JobRole = "Assistant Manager"
userRoles = Array("Access Administrator")
tabOption = "USERS"

addToUserInfo "DT_UserFirstName",userFirstName
addToUserInfo "DT_UserLastName", userLastName
addToUserInfo "DT_UserEmail", userEmailID
addToUserInfo "DT_UserCAID", caid
addToUserInfo "DT_UserRole" , userRole
addToUserInfo "DT_UserGroup", userGroup
addToUserInfo "DT_TabOption", tabOption

setCurrentView viewName
setCurrentModule moduleName
addToCache "DT_Location",locationtype
'*************************************************************
'schedulerLogin userName,password
'goToLoginView viewName
'goToModule moduleName
'selectBranch  branchType, branchName

addUser
'Function Name : AddUser
Public Function addUser()
	appendNum = RandomNumber(100,200)
	addToCache "DT_RandomNum",appendNum
	If Browser("Smart Lobby").Page("Smart Lobby").Frame("Frame_UserManager").Link("Add User").Exist(1) Then
		Select Case retrieveFromCache("DT_TabOption")
			Case "USERS"
				Browser("Smart Lobby").Page("Smart Lobby").Frame("Frame_UserManager").Link("Tab_Users").Click
			Case "ACCESSADMIN"
				Browser("Smart Lobby").Page("Smart Lobby").Frame("Frame_UserManager").Link("Tab_AccessAdmin").Click
		End Select
		Browser("Smart Lobby").Page("Smart Lobby").Frame("Frame_UserManager").Link("Add User").Click
		CustomSync Browser("AddEditUser").Page("AddEditUser"), False, "Add User page launched"
		'Enter User Information
		Browser("AddEditUser").Page("AddEditUser").WebEdit("UserFirstName").Set(userFirstName & appendNum)
		Browser("AddEditUser").Page("AddEditUser").WebEdit("UserLastName").Set(userFirstName & appendNum)
		Browser("AddEditUser").Page("AddEditUser").WebEdit("TxtEmail").Set userEmailID
		'Append Random number to CAID
		caid = caid	& appendNum
		Browser("AddEditUser").Page("AddEditUser").WebEdit("CAID").Set caid
		'List Box to select JobRole
		Browser("AddEditUser").Page("AddEditUser").WebEdit("txtJobRole").Click
		Browser("AddEditUser").Page("AddEditUser").WebList("List_JobRoles").Select jobRole
		Browser("AddEditUser").Page("AddEditUser").WebList("List_JobRoles").Click
		jobCode = Browser("AddEditUser").Page("AddEditUser").WebElement("JobCode").GetROProperty("innertext")
		addToCache "DT_JobCode", jobCode
		'Add all user gropus mentioned
		For Each group in userRoles
			Browser("AddEditUser").Page("AddEditUser").WebList("LstAvailableGroups").Select group
			Browser("AddEditUser").Page("AddEditUser").WebButton(">").Click
			count1 = Browser("AddEditUser").Page("AddEditUser").WebList("LstAssignedGroups").GetROProperty("items count")
			If  Browser("AddEditUser").Page("AddEditUser").WebList("LstAssignedGroups").GetItem(count1) = group Then
				logPass "USER GROUP - Selected User Group has been added successfully"
			End If
		Next
		Browser("AddEditUser").Page("AddEditUser").Link("SaveUser").Click
		'Check whether Add user has given any errors or not
		If Dialog("text:=Message from webpage").Exist(10) Then
			chdObj = Dialog("text:=Message from webpage").ChildObjects
			ERR_MSG = chdObj(2).GetRoProperty("text")				
			logFail "ADD USER DETAILS- " & "Err msg" & ERR_MSG
		ElseIf Browser("AddEditUser").Page("AddEditUser").WebElement("CAID currently exists.").Exist(1) Then
			logFail "ADD USER - CAID already exists"
		End If
		'Add User -Date Added Column verification Data
		currTime = FormatDateTime(Time(),4)
		Select Case(systemTimeZone)
			Case CENTRAL_TIME_ZONE
				DataTable.Value("DateAdded","USERDATA") = Date() &" " & Replace(currTime,Hour(currTime)&":",Hour(currTime)+1&":")				
			Case EASTERN_TIME_ZONE
				DataTable.Value("DateAdded","USERDATA") = expected_DateAdded			
			Case MOUNTAIN_TIME_ZONE
				DataTable.Value("DateAdded","USERDATA") = DateAdd("h",2,expected_DateAdded)			
			Case systemTimeZone = PACIFIC_TIME_ZONE
				DataTable.Value("DateAdded","USERDATA") = DateAdd("h",3,expected_DateAdded)
			Case Else
				logFail "System Time Zone - Not in US Format.Add sytem time zone format to select case for validation"
			End Select
	Else
		logFatal "Add User tab is not available - Check the application workflow"
	End If
End Function
'################################################
'Function Name : Check_And_AddUser
' # Purpose: To Add new user with the define role
'################################################
Public Function checkAndAddUser(userFirstName, userLastName)
   Browser("Smart Lobby").Page("Smart Lobby").Image("Users Manager").Click
   Set DESC_OBJ = Browser("Smart Lobby").Page("Smart Lobby").Frame("Frame_UserManager").WbfGrid("UserDataGrid")
	rCount = DESC_OBJ.RowCount
	cCount = DESC_OBJ.ColumnCount(1)
	isFound =False
	For i=2 to rCount
		record_Fullname = DESC_OBJ.GetCellData(i,1)
		If  record_Fullname = userFirstName & " " & userLastName Then
			isFound =True
			Exit For
		End If
	Next
	If isFound = False Then
		'To Do - Check workflows and add functionality accordingly
		Call addUser()
	End If
End Function

'#  FunctionName : Launch_EditUser
Public Function goToEditUser()
		If Browser("Smart Lobby").Page("Smart Lobby").Frame("Frame_UserManager").Link("Edit User").Exist(1) Then
			Browser("Smart Lobby").Page("Smart Lobby").Frame("Frame_UserManager").Link("Edit User").Click
			CustomeSync Browser("AddEditUser").Page("AddEditUser") , False, "Edit User page has been launched"
			'To check user has been selected or not
			If Dialog("text:=Message from webpage").Exist(1) Then
				logFail "Please select user to Edit - User has not been selected for Edit"
			Else
				Browser("AddEditUser").Page("AddEditUser").UDF_Init
			End If
			'To check whether Edit user page has been launched or not
			If Browser("AddEditUser").Page("AddEditUser").Exist(1) Then
				logPass "Edit User Page has been launched successfully"
			End If
		End If	
End Function

'#  FunctionName : EditUserDetails
Public Function EditUserDetails(userRole,userGroupArray())
		If Browser("AddEditUser").Page("AddEditUser").Exist(1) Then
		'Enter User Information
			Browser("AddEditUser").Page("AddEditUser").UDF_Init			
			Browser("AddEditUser").Page("AddEditUser").WebEdit("UserFirstName").UDF_SetValue DataTable.Value("EditUserFirstName", "USERDATA")
			Browser("AddEditUser").Page("AddEditUser").WebEdit("UserLastName").UDF_SetValue DataTable.Value("EditUserLastName","USERDATA")
			Browser("AddEditUser").Page("AddEditUser").WebEdit("TxtEmail").UDF_SetValue DataTable.Value("EditUserEmailID","USERDATA")
			Browser("AddEditUser").Page("AddEditUser").WebEdit("CAID").UDF_SetValue DataTable.Value("EditCAID","USERDATA")	   
		'List Box to select JobRole if not Empty
			If userRole<>"" Then
				Browser("AddEditUser").Page("AddEditUser").WebEdit("txtJobRole").Click
				Browser("AddEditUser").Page("AddEditUser").WebList("List_JobRoles").Select userRole
				Browser("AddEditUser").Page("AddEditUser").WebList("List_JobRoles").Click
				jobCode = Browser("AddEditUser").Page("AddEditUser").WebElement("JobCode").GetROProperty("innertext")
			End If
  			If userGroupArray <>"" Then
			'Add all user gropus mentioned
				For Each group in userGroupArray
					Browser("AddEditUser").Page("AddEditUser").WebList("LstAvailableGroups").Select group
					Browser("AddEditUser").Page("AddEditUser").WebButton(">").Click
					count1 = Browser("AddEditUser").Page("AddEditUser").WebList("LstAssignedGroups").GetROProperty("items count")
					If  Browser("AddEditUser").Page("AddEditUser").WebList("LstAssignedGroups").GetItem(count1) = group Then
						Reporter.ReportEvent micPass,"SELECTED USER GROUP", "Selected User Group has been added successfully"
					End If
				Next
			End If
			Browser("AddEditUser").Page("AddEditUser").Link("SaveUser").UDF_Click
		'Check whether edit user has given any errors or not
			If Dialog("text:=Message from webpage").Exist(10) Then
				chdObj = Dialog("text:=Message from webpage").ChildObjects
				ERR_MSG = chdObj(2).GetRoProperty("text")
				FUNC_EXE_STATUS = False
				Reporter.ReportEvent micFail, "EDIT USER DETAILS", "Err msg" & ERR_MSG
				Exit Function
			End If
		Else
			Reporter.ReportEvent micFail, "Edit User", "Edit User Button is either unavailable or not in enabled state"
		End If	
End Function

'#  FunctionName : EditUserLocation
Public Function EditUserLocation(setUserLocation)
   	Browser("AddEditUser").Page("AddEditUser").Link("Tab_Locations").UDF_Click
		isFound = False
		If Browser("AddEditUser").Page("AddEditUser").WebElement("innertext:="& setUserLocation,"index:=0").Exist(1) Then
			isFound = True
			Browser("AddEditUser").Page("AddEditUser").WebTable("text:="&setUserLocation).Highlight
			Browser("AddEditUser").Page("AddEditUser").WebTable("text:="&setUserLocation).WebCheckBox("index:=0").UDF_Click
		Else
			Set imgObj = Description.Create
			imgObj("micclass").Value = "Image"
			imgObj("html tag").Value = "IMG"
			Set chdImg = Browser("AddEditUser").Page("AddEditUser").ChildObjects(imgObj)
			Dim list1()
			Dim tempindex
			tempindex = 0
			For i=0 to chdImg.Count-1
					temp = chdImg(i).GetRoProperty("alt")
					If MID(temp,1,6) = "Expand" Then
						ReDim Preserve list1(tempindex)
						list1(tempindex) = temp
						tempindex = tempindex+1
					End If
			Next
			set myWebElement = Browser("AddEditUser").Page("AddEditUser").WebElement("innertext:=" & setUserLocation,"index:=0")
			For Each imgtext in list1
				'ADD EXCEPTIONAL STRINGS HERE-STRINGS WHICH CONTAIN SPECIAL CHARs
					If imgtext = "Expand MEAC CC (CC3 & VC3)  " Then
						imgtext = "Expand MEAC CC \(CC3 \& VC3\)  "
					End If
					Set myObj = Browser("AddEditUser").Page("AddEditUser").Image("alt:=" & imgtext)
					myObj.Click
					If myWebElement.Exist(1)Then
							Browser("AddEditUser").Page("AddEditUser").WebTable("text:="&setUserLocation).WebCheckBox("index:=0").UDF_Click
							isFound = True
							Exit For
					End If
					If Err.Number = -2147220989 Then
							Reporter.ReportEvent micFail, "Treeview contains duplicate Objects",imgtext
					End If
			Next
		End If
		If isFound Then
			Browser("AddEditUser").Page("AddEditUser").Link("SaveUser").UDF_Click
			If Dialog("text:=Message from webpage").Exist Then
				Set chdItems  = Dialog("text:=Message from webpage").ChildObjects
			End If
			Browser("Smart Lobby").Page("Smart Lobby").UDF_Init
			Reporter.ReportEvent micPass, "EditUser Location", "Location has been edited successfully"
		Else
			FUNC_EXE_STATUS = False
			Reporter.ReportEvent micFail, "Edit USer Location", "Requested Location not Found"
		End If
End Function

'#  FunctionName :  Navigate_UserManagerPageTabs
Public Function Navigate_UserManagerPageTabs(userManagerTabs)
		Select Case(userManagerTabs)
			Case USERMANAGER_USERS
				Browser("Smart Lobby").Page("Smart Lobby").Frame("Frame_UserManager").Link("Tab_Users").UDF_Click
			Case USERMANAGER_ACCESSADMIN
				Browser("Smart Lobby").Page("Smart Lobby").Frame("Frame_UserManager").Link("Tab_Users").UDF_Click
			End Select
End Function

'#  FunctionName : VerifyUserData
'######################################################################
Public Function VerifyUserData(userWorkFlow)
   Set DESC_OBJ = Browser("Smart Lobby").Page("Smart Lobby").Frame("Frame_UserManager").WbfGrid("UserDataGrid")
	   rCount = DESC_OBJ.RowCount
	   cCount = DESC_OBJ.ColumnCount(1)
		isFound =False
		Select Case(userWorkFlow)
			Case USER_WORKFLOW_ADDUSER
				For i=2 to rCount
					If UCASE(DESC_OBJ.GetCellData(i,1)) = UCASE(DataTable.Value("UserFirstName", "USERDATA") & " " & DataTable.Value("UserLastName","USERDATA")) AND _
						DateDiff("n", FormatDateTime(DESC_OBJ.GetCellData(i,3)),FormatDateTime(DataTable.Value("DateAdded","USERDATA"))) <5Then
						DESC_OBJ.Object.Rows(i-1).Click
						'Msgbox "Pass"
						isFound = True
						Exit For
					End If
				Next
			Case USER_WORKFLOW_EDITUSER
				For i=2 to rCount
					If UCASE(DESC_OBJ.GetCellData(i,1)) = UCASE(DataTable.Value("EditUserFirstName", "USERDATA") & " " & DataTable.Value("EditUserLastName","USERDATA")) Then
						DESC_OBJ.Object.Rows(i-1).Click
						'Msgbox "Pass"
						isFound = True
						Exit For
					End If
				Next
		End Select
		If isFound =False Then
			Reporter.ReportEvent micFail, "User Validation", "User Validation Failed"
		End If	
End Function
'##############################################################



