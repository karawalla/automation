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
userFirstName = DataTable.Value("UserFirstName",Global)
userLastName = DataTable.Value("UserLastName",Global)
userEmailID = DataTable.Value("UserEmailID",Global)
caid =DataTable.Value("Caid",Global)
JobRole = DataTable.Value("JobRole",Global)
userRole = DataTable.Value("UserRole",Global)
tabOption = DataTable.Value("TabOption",Global)

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
schedulerLogin userName,password
goToLoginView viewName
goToModule moduleName
'selectBranch  branchType, branchName

addUser
verifyUserData
goToEditUser
EditUserLocation("Arboretum")

'Function Name : AddUser
Public Function addUser()
	appendNum = RandomNumber(100,200)
	addToCache "DT_RandomNum",appendNum
	If Browser("Smart Lobby").Page("Smart Lobby").Frame("Frame_UserManager").Link("Add User").Exist(1) Then
		Select Case retrieveFromUserInfo("DT_TabOption")
			Case TAB_USERS
				Browser("Smart Lobby").Page("Smart Lobby").Frame("Frame_UserManager").Link("Tab_Users").Click
			Case TAB_ACCESSADMIN
				Browser("Smart Lobby").Page("Smart Lobby").Frame("Frame_UserManager").Link("Tab_AccessAdmin").Click
		End Select
		Browser("Smart Lobby").Page("Smart Lobby").Frame("Frame_UserManager").Link("Add User").Click
		CustomSync Browser("AddEditUser").Page("AddEditUser"), False, "Add User page launched"
		'Enter User Information
		Browser("AddEditUser").Page("AddEditUser").WebEdit("UserFirstName").Set userFirstName & appendNum
		Browser("AddEditUser").Page("AddEditUser").WebEdit("UserLastName").Set userLastName & appendNum
		Browser("AddEditUser").Page("AddEditUser").WebEdit("TxtEmail").Set userEmailID
		Browser("AddEditUser").Page("AddEditUser").WebEdit("CAID").Set caid	& appendNum

		'List Box to select JobRole
		Browser("AddEditUser").Page("AddEditUser").WebEdit("txtJobRole").Click
		Browser("AddEditUser").Page("AddEditUser").WebList("List_JobRoles").Select jobRole
		Browser("AddEditUser").Page("AddEditUser").WebList("List_JobRoles").Click
		jobCode = Browser("AddEditUser").Page("AddEditUser").WebElement("JobCode").GetROProperty("innertext")
		addToCache "DT_JobCode", jobCode

		If  userRole <>"" AND Not isEmpty(userRole)Then
			Browser("AddEditUser").Page("AddEditUser").WebList("LstAvailableGroups").Select userRole
			Browser("AddEditUser").Page("AddEditUser").WebButton(">").Click
			count1 = Browser("AddEditUser").Page("AddEditUser").WebList("LstAssignedGroups").GetROProperty("items count")
			If  Browser("AddEditUser").Page("AddEditUser").WebList("LstAssignedGroups").GetItem(count1) = userRole Then
				logPass "USER GROUP - Selected User Group has been added successfully"
			End If
			Else
				logFail "USER GROUP - User Group cannot be Empty.Please recheck input data"
		End If
		Browser("AddEditUser").Page("AddEditUser").Link("SaveUser").Click
		'Check whether Add user has given any errors or not
		If Dialog("text:=Message from webpage").Exist(1) Then
			myObj = Description.Create
			myObj("micclass").Value = "Static"
			chdObj = Dialog("text:=Message from webpage").ChildObjects(myObj)
			ERR_MSG = chdObj(1).GetRoProperty("text")	
			logFail "ADD USER DETAILS- " & "Err msg" & ERR_MSG
		ElseIf Browser("AddEditUser").Page("AddEditUser").WebElement("CAID currently exists.").Exist(1) Then
			logFail "ADD USER - CAID already exists"
		End If
		CustomSync Browser("Smart Lobby").Page("Smart Lobby"), False, "Add User page successfully closed"
	Else
		logFatal "Add User tab is not available - Check the application workflow"
	End If
End Function

Public Function verifyUserData()
   Set DESC_OBJ = Browser("Smart Lobby").Page("Smart Lobby").Frame("Frame_UserManager").WbfGrid("UserDataGrid")
	   rCount = DESC_OBJ.RowCount
	   cCount = DESC_OBJ.ColumnCount(1)
		isFound =False
		Select Case retrieveFromUserInfo("DT_TabOption")
			Case TAB_USERS
				For i=2 to rCount
					If UCase(DESC_OBJ.GetCellData(i,1)) = UCase(retrieveFromUserInfo("DT_UserFirstName")) & retrieveFromCache("DT_RandomNum") & _
						" " & UCase(retrieveFromUserInfo("DT_UserLastName")) & retrieveFromCache("DT_RandomNum") AND _
						UCase(DESC_OBJ.GetCellData(i,2)) = UCase(retrieveFromUserInfo("DT_UserEmail")) Then
							DESC_OBJ.Object.Rows(i-1).Click
							logPass "Add User -Verified and record has been created successfully"
							isFound = True
							Exit For
					End If
				Next
			Case TAB_ACCESSADMIN
				For i=2 to rCount
					If UCase(DESC_OBJ.GetCellData(i,1)) = UCase(retrieveFromUserInfo("DT_UserFirstName")) & " " & UCase(retrieveFromUserInfo("DT_UserLastName")) AND _
						UCase(DESC_OBJ.GetCellData(i,2)) = UCase(retrieveFromUserInfo("DT_UserCAID")) &retrieveFromCache("DT_RandomNum") AND _
						UCase(DESC_OBJ.GetCellData(i,3)) = UCase(retrieveFromUserInfo("DT_UserEmail"))Then
							DESC_OBJ.Object.Rows(i-1).Click
							logPass "Add User -Verified and record has been created successfully"
							isFound = True
							Exit For
					End If
				Next
		End Select				
		If isFound =False Then
			logFail "Add User -Verified and no record found"
		End If
End Function

Public Function goToEditUser()
		If Browser("Smart Lobby").Page("Smart Lobby").Frame("Frame_UserManager").Link("Edit User").Exist(1) Then
			Browser("Smart Lobby").Page("Smart Lobby").Frame("Frame_UserManager").Link("Edit User").Click
			'To check user has been selected or not
			If Dialog("text:=Message from webpage").Exist(1) Then
				logFail "Please select user to Edit - User has not been selected for Edit"
			Else
			CustomSync Browser("AddEditUser").Page("AddEditUser") , False, "Edit User page has been launched"
			End If
			'To check whether Edit user page has been launched or not
'			If Browser("AddEditUser").Page("AddEditUser").Exist(1) Then
'				logPass "Edit User Page has been launched successfully"
'			End If
		End If
End Function

''#  FunctionName : EditUserDetails
Public Function EditUserDetails(userRole,userGroupArray())
		If Browser("AddEditUser").Page("AddEditUser").Exist(1) Then
		'Enter User Information			
			Browser("AddEditUser").Page("AddEditUser").WebEdit("UserFirstName").Set "testedit"
			Browser("AddEditUser").Page("AddEditUser").WebEdit("UserLastName").Set "dataedit"
			Browser("AddEditUser").Page("AddEditUser").WebEdit("TxtEmail").Set "c@m.com"
			Browser("AddEditUser").Page("AddEditUser").WebEdit("CAID").Set "TESTCAID"
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
						logPass "SELECTED USER GROUP - Selected User Group has been added successfully"
					End If
				Next
			End If
			Browser("AddEditUser").Page("AddEditUser").Link("SaveUser").UDF_Click
		'Check whether edit user has given any errors or not
			If Dialog("text:=Message from webpage").Exist(10) Then
				chdObj = Dialog("text:=Message from webpage").ChildObjects
				ERR_MSG = chdObj(2).GetRoProperty("text")				
				logFail  "EDIT USER DETAILS" & ERR_MSG
				Exit Function
			End If
		Else
			logFail "Edit User - Edit User Button is either unavailable or not in enabled state"
		End If	
End Function

'#  FunctionName : EditUserLocation
Public Function EditUserLocation(setUserLocation)
   	Browser("AddEditUser").Page("AddEditUser").Link("Tab_Locations").Click
		isFound = False
		If Browser("AddEditUser").Page("AddEditUser").WebElement("innertext:="& setUserLocation,"index:=0").Exist(1) Then
			isFound = True			
			Browser("AddEditUser").Page("AddEditUser").WebTable("text:="&setUserLocation).WebCheckBox("index:=0").Click
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
							Browser("AddEditUser").Page("AddEditUser").WebTable("text:="&setUserLocation).WebCheckBox("index:=0").Click
							isFound = True
							Exit For
					End If
					If Err.Number = -2147220989 Then
							logFail "Treeview contains duplicate Objects" &imgtext
					End If
			Next
		End If
		If isFound Then
			Browser("AddEditUser").Page("AddEditUser").Link("SaveUser").Click
			If Dialog("text:=Message from webpage").Exist Then
				logFail "Failed edit user saving"
				'Set chdItems  = Dialog("text:=Message from webpage").ChildObjects
			Else
				CustomSync Browser("Smart Lobby").Page("Smart Lobby"), False, "Saved User Data"
				logPass "EditUser Location - Location has been edited successfully"
			End If
		Else
			logFail "Edit USer Location - Requested Location not Found"
		End If
End Function

Public Function searchInUserManager()
	Browser("Smart Lobby").Page("Smart Lobby").Image("Users Manager").Click
	Select Case retrieveFromCache("DT_TabOption")	
		Case TAB_USERS
			Browser("Smart Lobby").Page("Smart Lobby").Frame("Frame_UserManager").WebEdit("FName_Users").Set "Text1"
			Browser("Smart Lobby").Page("Smart Lobby").Frame("Frame_UserManager").WebEdit("LName_Users").Set "Text2"
		Case TAB_ACCESSADMIN
			Browser("Smart Lobby").Page("Smart Lobby").Frame("Frame_UserManager").Link("Tab_AccessAdmin").Click
			Browser("Smart Lobby").Page("Smart Lobby").Frame("Frame_UserManager").WebEdit("Email_AccessAdmin").Set "Text3"
			Browser("Smart Lobby").Page("Smart Lobby").Frame("Frame_UserManager").WebEdit("CAIID_AccessAdmin").Set 
			Browser("Smart Lobby").Page("Smart Lobby").Frame("Frame_UserManager").WebEdit("FirstName_AccessAdmin").Set
			Browser("Smart Lobby").Page("Smart Lobby").Frame("Frame_UserManager").WebEdit("LastName_AccessAdmin").Set	
	End Select
	Browser("Smart Lobby").Page("Smart Lobby").Frame("Frame_UserManager").WebButton("Search").Click
End Function

'################################################
'Function Name : Check_And_AddUser
' # Purpose: To Add new user with the define role
'################################################
'Public Function checkAndAddUser(userFirstName, userLastName)
'   Browser("Smart Lobby").Page("Smart Lobby").Image("Users Manager").Click
'   Set DESC_OBJ = Browser("Smart Lobby").Page("Smart Lobby").Frame("Frame_UserManager").WbfGrid("UserDataGrid")
'	rCount = DESC_OBJ.RowCount
'	cCount = DESC_OBJ.ColumnCount(1)
'	isFound =False
'	For i=2 to rCount
'		record_Fullname = DESC_OBJ.GetCellData(i,1)
'		If  record_Fullname = userFirstName & " " & userLastName Then
'			isFound =True
'			Exit For
'		End If
'	Next
'	If isFound = False Then
'		'To Do - Check workflows and add functionality accordingly
'		Call addUser()
'	End If
'End Function
'


''#  FunctionName : VerifyUserData
''######################################################################
