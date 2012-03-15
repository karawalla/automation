'@@ 
'@ Name: launchScheduler
'@ Description: Utility function to launch the scheduler application.
'@ Useful to launch Scheduler applications of all environements.
'@ Arg1: url Value of the argument which needs to be launched
'@ Return: None
'@ Example: launchScheduler "http://qaweb2.ncr.com/SSMPortalBOA_Trunk/Login.aspx"
'@ History:
'@ Tags:
'@@
Public Function launchScheduler(Byval url)
	WebBrowserInvoke "IE", url
	CustomSync Browser("Smart Lobby-LoginPage").Page("Smart Lobby-LoginPage"), False, "Navigated to Login Page - "
	setAppWorkFlow  APP_WORKFLOW_LOGIN
End Function

'@@ 
'@ Name: schedulerLogin
'@ Description: Utility function to login to the scheduler application
'@ Arg1: userName Value to be used as User ID for login
'@ Arg2: password Value to be used as Password for login
'@ Return: None
'@ Example: schedulerLogin("userName", "password")
'@ History:
'@ Tags:
'@@
Public Function schedulerLogin(Byval userName, Byval password)
	'CustomSync Browser("Smart Lobby-LoginPage").Page("Smart Lobby-LoginPage").Sync, False, "Login Page Launched"
	Browser("Smart Lobby-LoginPage").Page("Smart Lobby-LoginPage").WebEdit("UserName").Set userName
	Browser("Smart Lobby-LoginPage").Page("Smart Lobby-LoginPage").WebEdit("Password").Set password
	Browser("Smart Lobby-LoginPage").Page("Smart Lobby-LoginPage").Image("LoginBtn").Click
	'To check
    If Browser("Smart Lobby-LoginPage").Dialog("Windows Internet Explorer").Exist(1) Then
		Browser("Smart Lobby-LoginPage").Dialog("Windows Internet Explorer").Activate
		Browser("Smart Lobby-LoginPage").Dialog("Windows Internet Explorer").WinButton("Yes").Highlight
		Browser("Smart Lobby-LoginPage").Dialog("Windows Internet Explorer").WinButton("Yes").Click
	End If
	SetValidationObject Browser("SmartLobby").Page("SmartLobby").Image("ActivityMonitoring")
	CustomSync Browser("SmartLobby").Page("SmartLobby"),False, "Smart Lobby Home page launched successfully"			
End Function

'@@ 
'@ Name: schedulerLogout
'@ Description: Utility function to logout to the scheduler application
'@ Return: None
'@ Example: schedulerLogout()
'@ History:
'@ Tags:
'@@
Public Function schedulerLogout()
	If  Browser("SmartLobby").Page("SmartLobby").Exist(1)Then
		Browser("SmartLobby").Page("SmartLobby").Image("Log Out").Click
		wait(2)
		Browser("title:=Bank of America \| Simplified Sign-On \| Logoff").Highlight
		Browser("title:=Bank of America \| Simplified Sign-On \| Logoff").Close
		logPass "LOGOUT - Logged out and closed the application"
	Else
		logFail "LOGOUT - Home page not getting displayed for Logging out"
	End If
End Function

'@@ 
'@ Name: goToLoginView
'@ Description: Utility function to click on Login View during the first time launch of scheduler application
'@ Arg1: viewName Value to be selected - Either VIEW_PLATFORM or VIEW_LOBBY
'@ Return: None
'@ Example: goToLoginView("viewName")
'@ History:
'@ Tags:
'@@
Public Function goToLoginView(Byval viewName)
    Browser("SmartLobby").Page("SmartLobby").Image("ActivityMonitoring").Click
	SetValidationObject Browser("SmartLobby").Page("SmartLobby").Frame("Frame_User Permissions")
	CustomSync Browser("SmartLobby").Page("SmartLobby"), False, "Navigated to Login View Selection Page."
	logInfo "Selecting View:" & viewName
    If viewName = VIEW_PLATFORM Then
        Browser("SmartLobby").Page("SmartLobby").Frame("Frame_User Permissions").Image("imgPlatform").Click
    Else
        Browser("SmartLobby").Page("SmartLobby").Frame("Frame_User Permissions").Image("imgLobby").Click	    
	End If
	SetValidationObject Browser("SmartLobby").Page("SmartLobby").Frame("Frame_ActivityManager")
    CustomSync Browser("SmartLobby").Page("SmartLobby"), False, "Navigated to " & viewName
	'setCurrentView viewName
End Function

'@@ 
'@ Name: goToModule
'@ Description: Utility function to click on Module in the Scheduler Application
'@ Arg1: moduleName to be selected in the Scheduler application
'@ Return: None
'@ Example: goToModule("moduleName")
'@ History:
'@ Tags:
'@@
Public Function goToModule(Byval moduleName)
	Select Case moduleName
		Case MODULE_ACTIVITYMONITORING
			Browser("SmartLobby").Page("SmartLobby").Image("ActivityMonitoring").Click			
			CustomFrameSync Browser("SmartLobby").Page("SmartLobby").Frame("Frame_ActivityManager"), False, "Navigated to Activity Monitoring Page"
		Case MODULE_ENTERPRISEMANAGER
			Browser("SmartLobby").Page("SmartLobby").Image("EnterpriseManager").Click
			CustomFrameSync Browser("SmartLobby").Page("SmartLobby").Frame("Frame_EnterpriseManager"), False, "Navigated to Enterprise Manager Page"
		Case MODULE_APPOINTMENTMANAGER
			Browser("SmartLobby").Page("SmartLobby").Image("AppointmentManager").Click
			CustomFrameSync Browser("SmartLobby").Page("SmartLobby").Frame("Frame_AppointmentManager"), False, "Navigated to Appointment Manager Page"
		Case MODULE_USERMANAGER
			Browser("SmartLobby").Page("SmartLobby").Image("Users Manager").Click
			CustomFrameSync Browser("SmartLobby").Page("SmartLobby").Frame("Frame_UserManager"), False, "Navigated to User Manager Page"
			'Case MODULE_REPORTMANAGER
			'Browser("Smart Lobby").Page("Smart Lobby").Image("Reports Manager").Click
			'Add CustomSync
		Case Else
			logFatal "Given " & moduleName & "  not Available. Please Check"
		End Select
		setCurrentModule moduleName
End Function

'@@ 
'@ Name: selectBranch
'@ Description: Utility function to select desired branch from the tree View in the Scheduler Application
'@ Arg1: branchType to be selected like Stype or Ytype
'@ Arg2: branchName to be selected
'@ Return: None
'@ Example: selectBranch("stype","Arborteum")
'@ History:
'@ Tags:
'@@
Public Function selectBranch(ByVal branchType,ByVal branchName)
		'To check and add validations
		Select Case branchType
			Case BRANCH_TYPE_S
			Case BRANCH_TYPE_Y
		End Select
		'Click on SHOW LOCATOR
		If Browser("SmartLobby").Page("SmartLobby").WebElement("ShowLocator").Exist(1) Then
			Browser("SmartLobby").Page("SmartLobby").WebElement("ShowHideLocatorArrow").Click
		End If
		'VERIFY WEB ELEMENT TOGGLED TO HIDE LOCATOR
		If Browser("SmartLobby").Page("SmartLobby").WebElement("HideLocator").Exist(1) Then
			logPass "Web Element toggled to HIDE LOCATOR"
		End If
        wait(1)
		'Select BANK LOCATION
		Set imgObj = Description.Create
		imgObj("micclass").Value = "Image"
		imgObj("html tag").Value = "IMG"
		Set chdImg = Browser("SmartLobby").Page("SmartLobby").Frame("BankLocation_TreeView").ChildObjects(imgObj)
		Dim list1()
		Dim tempindex
		tempindex = 0
		locType = retrieveFromCache("DT_Location")
		If locType ="CC" OR locType ="VC"Then
			For i=0 to chdImg.Count-1
				temp = chdImg(i).GetRoProperty("alt")
				If MID(temp,1,11) = "Expand Call" OR MID(temp,1,12) = "Expand Video"Then
					ReDim Preserve list1(tempindex)
					list1(tempindex) = temp
					tempindex = tempindex+1
				End If
			Next
		Else
			For i=0 to chdImg.Count-1
				temp = chdImg(i).GetRoProperty("alt")
				If MID(temp,1,6) = "Expand" Then
					ReDim Preserve list1(tempindex)
					list1(tempindex) = temp
					tempindex = tempindex+1
				End If
			Next
		End If
		set myLink = Browser("SmartLobby").Page("SmartLobby").Frame("BankLocation_TreeView").Link("text:=" & branchName)
		isFound = False
		For Each imgtext in list1
		'ADD EXCEPTIONAL STRINGS HERE-STRINGS WHICH CONTAIN SPECIAL CHARs
			If imgtext = "Expand MEAC CC (CC3 & VC3)  " Then
				imgtext = "Expand MEAC CC \(CC3 \& VC3\)  "
			End If
			Set myObj = Browser("SmartLobby").Page("SmartLobby").Frame("BankLocation_TreeView").Image("alt:=" & imgtext)
			myObj.Click
			If myLink.Exist(1)Then
				myLink.Click
				isFound = True
				Exit For
			End If
		Next
		If isFound = False AND tempindex <>0 Then
			logFatal "Select Bank Loc through Tree View Failed "
		End If
End Function

'@@ 
'@ Name: goToScheduleAppointment
'@ Description: Utility function to click on Schedule Appoinment button in Scheduler Appointment
'@ Return: None
'@ Example: goToScheduleAppointment()
'@ History:
'@ Tags:
'@@
Public Function goToScheduleAppointment()
		Select Case CURRENT_MODULE
			CASE MODULE_ACTIVITYMONITORING					
					Browser("SmartLobby").Page("SmartLobby").Frame("Frame_ActivityManager").WebElement("Schedule").Click                    
			CASE MODULE_APPOINTMENTMANAGER
					Browser("SmartLobby").Page("SmartLobby").Frame("Frame_AppointmentManager").Link("Schedule").Click				
		End Select
		CustomSync Browser("SelectBankTopic_Step2").Page("SelectBankTopic_Step2"), False, "Launched Select Discussion Topic Page"		
'		setCurrentWorkFlow WORKFLOW_SCHEDULE
End Function

'@@ 
'@ Name: selectTopicforDiscussion
'@ Description: Utility function to select Discussion topics and fill the details in Discussion Topic page
'@ Arg1: dictInData Dictionary containing Discussion Topics,account type to be selected, comments and Language fields
'@ Return: None
'@ Example: selectTopicforDiscussion(dictionaryObj)
'@ History:
'@ Tags:
'@@
Public Function selectTopicforDiscussion(dictInData)
   'Step2
	If Browser("SelectBankTopic_Step2").Page("SelectBankTopic_Step2").Exist(10) Then
		'Browser("Step2:Select a Topic for Bank").Page("Step2:Select a Topic for Bank").Highlight
		'Select Account Type
		Select Case dictInData.Item("DT_AccType")
			Case ACCTYPE_PB
				Browser("SelectBankTopic_Step2").Page("SelectBankTopic_Step2").WebRadioGroup("BankingType").Select("PB")
				'AddToCache "Role","PB"				 
			Case ACCTYPE_SBB
				Browser("SelectBankTopic_Step2").Page("SelectBankTopic_Step2").WebRadioGroup("BankingType").Select("SBB")
				Browser("SelectBankTopic_Step2").Page("SelectBankTopic_Step2").WebRadioGroup("empno").Select("Yes")
				'AddToCache  "Role","SBB"	
			Case ACCTYPE_PPB
				'AddToCache "Role","PPB"
			Case ACCTYPE_IME
				Browser("SelectBankTopic_Step2").Page("SelectBankTopic_Step2").WebRadioGroup("InvestmentWithMerillEdge").Select("IME")
			Case Else
           		LogFatal "Selected Bank Option Not Availble for Branch"				
			End Select
			'Select Discussion Topics			
			selectTopic = Array(dictInData.Item("DT_Topic1"),dictInData.Item("DT_Topic2"))
			'Browser("Select a Topic for Bank").Page("Select a Topic for Bank").Init
			For Each topic in selectTopic
				On Error Resume Next
				If topic <>"" Then
					tempArr = Split(topic,":")
					Set Desc= Description.Create
					Desc("micclass").value = "WebTable"
					Desc("innertext").value = tempArr(0)
					selectCount =0
					set matchingTables = Browser("SelectBankTopic_Step2").Page("SelectBankTopic_Step2").ChildObjects(Desc)
					'If matchingTables.Count >0 Then
					Set chkBox = matchingTables(0).ChildItem(1,1,"WebCheckBox", 0)
                    chkBox.Set "ON"
					wait 2
					Set Desc= Description.Create
					Desc("micclass").value = "WebTable"
					set subTables =    matchingTables(0).ChildObjects(Desc)
					For i = 1 to subTables(0).RowCount
						For j=1 to UBound(tempArr)
							subTopic = tempArr(j)
							If  subTables(0).GetCellData(i,1) = subTopic Then
								Set chkBox = subTables(0).ChildItem(i,1,"WebCheckBox", 0)						
								chkBox.Set "ON"
								'selectCount =selectCount+1
							End If
						Next
					Next
				End If
				'End If
			Next
			If Err.Number <> 0 Then
				Set DescObj= Description.Create
				DescObj("micclass").value = "Webcheckbox"
				Set setAnyObject = Browser("SelectBankTopic_Step2").Page("SelectBankTopic_Step2").ChildObjects(DescObj)
				SelectRandom = RandomNumber(1,setAnyObject.Count)
				ObjToSelect = setAnyObject(SelectRandom).GetRoProperty("name")
				Set chkBox = Browser("SelectBankTopic_Step2").Page("SelectBankTopic_Step2").WebCheckBox("name:="&ObjToSelect)
				'chkBox.Highlight
				chkBox.Set "ON"
				logFail "Given Discussion Topic not available in Application.Selected Topic:" & ObjToSelect &"and continued execution of workflow"
			End If
			'Enter Comments and Language
			Browser("SelectBankTopic_Step2").Page("SelectBankTopic_Step2").WebEdit("comments").Set dictInData.Item("DT_Comments")
			If dictInData.Item("DT_AccType") <> ACCTYPE_PPB Then
				Browser("SelectBankTopic_Step2").Page("SelectBankTopic_Step2").WebEdit("Language").Set dictInData.Item("DT_Language")
			End If
			Select Case CURRENT_WORKFLOW
				Case WORKFLOW_WALKIN
					Browser("SelectBankTopic_Step2").Page("SelectBankTopic_Step2").WebElement("Submit").Click
					CustomSync Browser("SmartLobby").Page("SmartLobby"), False, "Clicked on Submit Button"
				Case Else 'Case WORKFLOW_SCHEDULE
					Browser("SelectBankTopic_Step2").Page("SelectBankTopic_Step2").Link("Continue").Click
					CustomSync Browser("SelectDateAndTime_Step3").Page("SelectDateAndTime_Step3"),False,"Clicked on Continue button"
			End Select
	Else
		logFatal "Select a Topic for Bank - Page Not Found"	
	End If
End Function

'@@ 
'@ Name: selectDateTime
'@ Description: Utility function to select Discussion topics and fill the details in Discussion Topic page
'@ Arg1: zipCode to set the zipcode edit box value
'@ Arg2: dateToSelect criteria based on whcih date will be selected for appointment
'@ Possible options for dateToSelect - DATESELECT_TODAY,DATESELECT_TOMORROW,DATESELECT_DAYAFTERTOMORROW,DATESELECT_RANDOM
'@ Return: None
'@ Example: selectDateTime("75038",DATESELECT_RANDOM) - To select any date in the calendar
'@ History:
'@ Tags:
'@@
Public Function selectDateTime(ByVal zipCode, ByVal dateToSelect)
   'STEP 3
	If Browser("SelectDateAndTime_Step3").Page("SelectDateAndTime_Step3").Exist(1) Then
		If APP_WORKFLOW = APP_WORKFLOW_LOGIN Then
			'Set and Validate Zip Code
			Call validateCalendarTimeZones(zipCode)
		End If
		'Select Date Time in Calendar
		If Browser("SelectDateAndTime_Step3").Page("SelectDateAndTime_Step3").WebTable("WebTable_SelectDate").Exist(1) Then
			Set DESC_OBJ = Browser("SelectDateAndTime_Step3").Page("SelectDateAndTime_Step3").WebTable("WebTable_SelectDate")
			'Useful when application displays calender after 2 weeks
			If dateToSelect = DATESELECT_RANDOM Then
				Call dateSelectionRandom(DESC_OBJ)
				'Useful when application displays current week calender and need to select Specific date and available time
			ElseIf dateToSelect = DATESELECT_TODAY OR dateToSelect = DATESELECT_TOMORROW OR _
				dateToSelect = DATESELECT_DAYAFTERTOMORROW Then
				Call dateSelectionSpecificDate(DESC_OBJ,dateToSelect)
			End If
			Browser("SelectDateAndTime_Step3").Page("SelectDateAndTime_Step3").WebElement("Continue").Click
			CustomSync Browser("ContactInfoForBank_Step4").Page("ContactInfoForBank_Step4"), "False", "Clicked on continue button on Date selection page."
		Else
			logFatal "SelectDateTime", "Failed"
		End If
	Else
		LogFatal "Step3:Select a Day and Time page not Found"
	End If
End Function

'@@ 
'@ Name: validateCalendarTimeZones
'@Description : It is a supporting function for selectdateTime function and not used for calling explicitly
'@To Validate the default existing timezone calendar with the user entered zipcode timezone calendar
'@If by default zipcode field value is <blank> validates skips and enteres the new user entered timezone.
'@Reports whether validation is Pass or Fail
'@ Arg1: zipCode -It will passed from selectdateTime function.
'@ Return: None
'@ History:
'@ Tags:
'@@
Public Function validateCalendarTimeZones(zipCode)
	If Browser("SelectDateAndTime_Step3").Page("SelectDateAndTime_Step3").Exist(1) Then
		'Check zipCode parameter
		If Not isEmpty(zipCode) Then
			'Check for Zipcode edit box -Empty or Not
			If Browser("SelectDateAndTime_Step3").Page("SelectDateAndTime_Step3").WebEdit("zipcode").GetROProperty("value") <>"" Then
				currApp_TimeZone = Browser("SelectDateAndTime_Step3").Page("SelectDateAndTime_Step3").WebElement("TIMEZONE").GetROProperty("innertext")
               	curr_StartTime = HOUR(Browser("SelectDateAndTime_Step3").Page("SelectDateAndTime_Step3").WebTable("WebTable_SelectDate").GetCellData(3,1))
				'Set Differnt Zone Time values for Validation
				Select Case Trim(currApp_TimeZone)
					Case PACIFIC_TIME_ZONE
						pacificTime = curr_StartTime
						mountainTime = curr_StartTime +1
						centralTime = curr_StartTime +2
						easternTime = curr_StartTime +3
					Case MOUNTAIN_TIME_ZONE
						pacificTime = curr_StartTime -1
						mountainTime = curr_StartTime
						centralTime = curr_StartTime +1
						easternTime = curr_StartTime +2
					Case CENTRAL_TIME_ZONE
						pacificTime = curr_StartTime -2
						mountainTime = curr_StartTime-1
						centralTime = curr_StartTime
						easternTime = curr_StartTime +1
					Case EASTERN_TIME_ZONE
						pacificTime = curr_StartTime -3
						mountainTime = curr_StartTime-2
						centralTime = curr_StartTime -1
						easternTime = curr_StartTime
				End Select
			End If
		End If
		'Set zipCode Value and click on Display Calender
		Browser("SelectDateAndTime_Step3").Page("SelectDateAndTime_Step3").WebEdit("zipcode").Set zipCode
		Browser("SelectDateAndTime_Step3").Page("SelectDateAndTime_Step3").WebElement("Display Calendar").Click				
		Browser("SelectDateAndTime_Step3").Page("SelectDateAndTime_Step3").WebTable("WebTable_SelectDate").Init
		'Check Calender Web table has been refreshed or not
		calenderRefresh = True
		Browser("SelectDateAndTime_Step3").WinStatusBar("StatusbarMsg").DblClick 10,10
		If Browser("SelectDateAndTime_Step3").Window("WindowsInternetExplorer").Exist(1) Then
			logFail "Error on Page:Calender has not been refreshed to new timeZone"
			calenderRefresh = False
			Browser("SelectDateAndTime_Step3").Window("WindowsInternetExplorer").Close
		End If
		'Validate new Times based on new TimeZone if calender webtable has refreshed
		If Not isEmpty(currApp_TimeZone) AND  calenderRefresh = True Then
			new_TimeZone = Browser("SelectDateAndTime_Step3").Page("SelectDateAndTime_Step3").WebElement("TIMEZONE").GetROProperty("innertext")
			new_StartTime = HOUR(Browser("SelectDateAndTime_Step3").Page("SelectDateAndTime_Step3").WebTable("WebTable_SelectDate").GetCellData(3,1))
			Select Case Trim(new_TimeZone)
				Case PACIFIC_TIME_ZONE
					verifyValue new_StartTime,pacificTime,"Calendar time updated to new time Zone"
				Case MOUNTAIN_TIME_ZONE					
					verifyValue new_StartTime,pacificTime,"Calendar time updated to new time Zone"			
				Case CENTRAL_TIME_ZONE		
					verifyValue new_StartTime,centralTime,"Calendar time updated to new time Zone"
				Case EASTERN_TIME_ZONE
					verifyValue new_StartTime,easternTime,"Calendar time updated to new time Zone"
			End Select
		End If
	End If
End Function

'@@ 
'@ Name: dateSelectionRandom
'@Description : It is a supporting function for selectdateTime function and not used for calling explicitly
'@Utility function to select random Date and time in calendar
'@Returns selected Date and Time selected in the calendar
'@ Arg1: DESC_OBJ-It will passed from selectdateTime function.
'@ Return: None
'@ History:
'@ Tags:
'@@
Public Function dateSelectionRandom(DESC_OBJ)
	Dim iArray()
	Set objstatic=description.Create
	objstatic("html tag").Value="A"
	objstatic("micclass").Value = "Link"
	chkCount =0
	For  chkCount =0 to 2
		Set FramesAll = DESC_OBJ.ChildObjects(objstatic)
        'Check if text is not empty and then add if not Empty
		indx=0
		For i=1 to FramesAll.Count-1
			If FramesAll(i).Object.tabIndex <> -1 Then
				ReDim Preserve iArray(indx)
				iArray(indx) = FramesAll(i).GetROProperty("text")
				indx=indx+1
			End If
		Next
		'If current week has no slots
		'If  isEmpty(UBOUND(iArray)) Then
		If indx =0 Then		
			logWarning "No slot available for Current Week"
			Browser("SelectDateAndTime_Step3").Page("SelectDateAndTime_Step3").Image("Move to next week").Click				
			DESC_OBJ.Init
			AssertObjExists DESC_OBJ, "Waited for Calender Refresh after clicked on Move to next week "
		Else
			Exit For
		End If
	Next
	If  chkCount < 2Then
		randomSelect = False
		While(randomSelect = False)
			sel_Randomtime =RandomNumber.Value(0,UBOUND(iArray))
			myTime = "text:=" + iArray(sel_Randomtime)
			sel_RandomIndex = RandomNumber.Value(0,4)
			indx = "index:="	& sel_RandomIndex
		'Check whether random date and time objects or not
			If Browser("SelectDateAndTime_Step3").Page("SelectDateAndTime_Step3").Link(myTime,indx).Exist(1) Then
				'Browser("Step3:Select a Day and Time").Page("Step3:Select a Day and Time").Link(myTime,indx).Highlight
				Browser("SelectDateAndTime_Step3").Page("SelectDateAndTime_Step3").Link(myTime,indx).Click
				randomSelect = True
				'Pathname object value contains both Date and Time
				temp = Browser("SelectDateAndTime_Step3").Page("SelectDateAndTime_Step3").Link(myTime,indx).Object.pathname
				'Set Date Time format to match with the application
				myarr = Split(temp,"'")
				DatenTime = myarr(1)
				DatenTime = Split(DatenTime," ",2)
				selectedDay = DatenTime(0)
				selectedDay = Replace(selectedDay,"/","-")
				selectedTime = DatenTime(1)
				'Add Date and Time to RunTime Cache for Validation
				addToCache "DT_ApptDate", selectedDay
				addToCache "DT_ApptTime",selectedTime
				If  Browser("SelectDateAndTime_Step3").Page("SelectDateAndTime_Step3").WebElement("TIMEZONE").Exist(1) Then
				'To Do
					'DataTable.Value("selectioninTimeZone","TESTDATA") = Browser("Step3:Select a Day and Time").Page("Step3:Select a Day and Time").WebElement("TIMEZONE").GetROProperty("innertext")
				End If
			End If
		Wend
	End If
End Function

'@@ 
'@ Name: dateSelectionSpecificDate
'@Description : It is a supporting function for selectdateTime function and not used for calling explicitly
'@Utility function to select Specific date and any time based on the user input in calendar
'@Returns selected Date and Time selected in the calendar
'@ Arg1: DESC_OBJ-It will passed from selectdateTime function.
'@ Return: None
'@ History:
'@ Tags:
'@@
Public Function dateSelectionSpecificDate(DESC_OBJ,dateToSelect)
	'Traverse thru the Row2 all columns to get the Date
	totalColumns = DESC_OBJ.ColumnCount(3)
	dateFound = False
	colIndex = 0
	For i=2 to totalColumns
		columnDate= DESC_OBJ.GetCellData(2,i)
		temp = Split(columnDate," ")
		columnDate = temp(1)
		'Get System Current Date
		currDate = Day(Date)
		'Set the column to Today Date
		If  currDate= cInt(columnDate) Then
			colIndex = i
			logPass "Current Calender displayed has today's date option"
			dateFound = True
			Exit For
		End If
	Next
	'If Date not Found
	If dateFound = False Then
		logFatal "Currently displayed calender does not contain today's Date"	
	End If
	'Set ColIndex based on date selection option
	If colIndex <>0 Then
		Select Case dateToSelect
			Case DATESELECT_TODAY
				colIndex = colIndex
			Case DATESELECT_TOMORROW
				colIndex = colIndex + 1
				If colIndex >totalColumns Then
					Browser("SelectDateAndTime_Step3").Page("SelectDateAndTime_Step3").Image("Move to next week").Click
					colIndex = colIndex - totalColumns +1
				End If
			Case DATESELECT_DAYAFTERTOMORROW
				colIndex = colIndex + 2
				If colIndex > totalColumns Then
					Browser("SelectDateAndTime_Step3").Page("SelectDateAndTime_Step3").Image("Move to next week").Click
					colIndex = colIndex - totalColumns +1
				End If
			End Select
		End If
		'Get the Total Rows in DataTable
		rCount = Browser("SelectDateAndTime_Step3").Page("SelectDateAndTime_Step3").WebTable("WebTable_SelectDate").RowCount
		'Traverse through the current date column to check the available slots
		slotFound = False
		For i=3 to rCount -1
			If DESC_OBJ.GetCellData(i,colIndex) <>"Unavailable " Then
				set myLink=DESC_OBJ.ChildItem(i,colIndex,"Link",0)
				'myLink.Highlight
				'Get Return Values -Selected Date and Time
				If dateToSelect = DATESELECT_TODAY Then
					selectedDay = Date
				ElseIf dateToSelect = DATESELECT_TOMORROW Then
					selectedDay = Date +1
				ElseIf dateToSelect = DATESELECT_DAYAFTERTOMORROW Then
					selectedDay = Date +2
				Else
					logFatal "Invalid Date Selection.Please check input"
				End If				
				selectedDay = Replace(selectedDay,"/","-")
				selectedTime = myLink.GetROProperty("innertext")				
				myLink.Click
				slotFound = True
				'Write date and time values to RunTime Cache for Validation				
				addToCache "DT_ApptDate", selectedDay
				addToCache "DT_ApptTime",selectedTime				
				'If  Browser("Step3:Select a Day and Time").Page("Step3:Select a Day and Time").WebElement("TIMEZONE").Exist(1) Then
					'DataTable.Value("selectioninTimeZone","TESTDATA") = Browser("Step3:Select a Day and Time").Page("Step3:Select a Day and Time").WebElement("TIMEZONE").GetROProperty("innertext")
				'End If
				logPass "Appt slot has been selected"
				Exit For
			End If
		Next
		'If Slot not Found
		If slotFound = False Then
			logFatal "Appointment slots are not available for" & dateToSelect			
		End If
End Function

'@@ 
'@ Name: provideContactInformation
'@ Description: Utility function to fill the user information in Provide Contact Information page
'@ Arg1: dictInData Dictionary object which contains user information like lastName,firstName,email etc
'@ Example: provideContactInformation(dictionayObj)
'@ Return:None
'@ History:
'@ Tags:
'@@
Public Function provideContactInformation(dictInData)
	'Step 4
	'To Append Index to Last and FirstNames for unique data set
	If Browser("ContactInfoForBank_Step4").Page("ContactInfoForBank_Step4").Exist(10) Then
		CountLastFirstName = RandomNumber(1,100)
		addToCache "DT_CountLastFirstName",CountLastFirstName		
		Browser("ContactInfoForBank_Step4").Page("ContactInfoForBank_Step4").WebEdit("FirstName").Set dictInData.Item("DT_FirstName") & CountLastFirstName
		Browser("ContactInfoForBank_Step4").Page("ContactInfoForBank_Step4").WebEdit("LastName").Set dictInData.Item("DT_LastName") & CountLastFirstName
		Browser("ContactInfoForBank_Step4").Page("ContactInfoForBank_Step4").WebEdit("Email").Set dictInData.Item("DT_Email")
		Browser("ContactInfoForBank_Step4").Page("ContactInfoForBank_Step4").WebEdit("ReEnterEmail").Set dictInData.Item("DT_ReEnterEmail")
		Browser("ContactInfoForBank_Step4").Page("ContactInfoForBank_Step4").WebEdit("ContactNumber").Set dictInData.Item("DT_ContactNumber")
		'Set Existing Customer
		If  dictInData.Item("DT_AccType")<> ACCTYPE_PPB Then
			Browser("ContactInfoForBank_Step4").Page("ContactInfoForBank_Step4").WebRadioGroup("ExistingCustomer").Select dictInData.Item("DT_ExistingCustomer")
		End If
		'Set check Box Value -Both
		If  UCASE(dictInData.Item("DT_Reminder_Both"))="ON"Then
			Browser("ContactInfoForBank_Step4").Page("ContactInfoForBank_Step4").WebCheckBox("Both").Set "ON"
			If Browser("ContactInfoForBank_Step4").Page("ContactInfoForBank_Step4").WebCheckBox("PhoneReminder").GetROProperty("checked") =1 And _
				Browser("ContactInfoForBank_Step4").Page("ContactInfoForBank_Step4").WebCheckBox("TxtMessgaetoMobile").GetROProperty("checked") =1 Then
				logPass "Reminder Options - Checking both has checked both phone reminder and text message options"
			Else
				logFail "Reminder Options - Checking both has not checked both phone reminder and text message options"
			End If
				Browser("ContactInfoForBank_Step4").Page("ContactInfoForBank_Step4").WebEdit("TextMsgMobilePhone").Set dictInData.Item("DT_ContactNumber")
				If Browser("ContactInfoForBank_Step4").Page("ContactInfoForBank_Step4").WebEdit("PhoneReminderContactNumber").GetROProperty("value") <> "" Then
					logPass "PHONE REMINDER FIELD - Phone reminder field auto populated"
				Else
					logFail "PHONE REMINDER FIELD - Phone reminder field not auto populated"
				End If
			End If
			'Set check Box Value - Text Message Only
			If  UCASE(dictInData.Item("DT_Reminder_TxtMsg")) ="ON"Then
				Browser("ContactInfoForBank_Step4").Page("ContactInfoForBank_Step4").WebCheckBox("TxtMessgaetoMobile").Set "ON"
				Browser("ContactInfoForBank_Step4").Page("ContactInfoForBank_Step4").WebEdit("TextMsgMobilePhone").Set dictInData.Item("DT_ContactNumber")
			End If
			'Set check Box Value - Phone Reminder Only
			If  UCASE(dictInData.Item("DT_Reminder_Phone")) ="ON"Then
				Browser("ContactInfoForBank_Step4").Page("ContactInfoForBank_Step4").WebCheckBox("PhoneReminder").Set "ON"
				If Browser("ContactInfoForBank_Step4").Page("ContactInfoForBank_Step4").WebEdit("PhoneReminderContactNumber").GetROProperty("value") = dictInData.Item("DT_ContactNumber")Then
					logPass "PHONE REMINDER FIELD - Phone reminder field auto populated"
				Else
					logFail "PHONE REMINDER FIELD - Phone reminder field not auto populated"
			End If
		End If
		'Check if both ph reminder and text reminder checkboxes are On, both should be ON by default
		If Browser("ContactInfoForBank_Step4").Page("ContactInfoForBank_Step4").WebCheckBox("PhoneReminder").GetROProperty("checked") =1 And _
			Browser("ContactInfoForBank_Step4").Page("ContactInfoForBank_Step4").WebCheckBox("TxtMessgaetoMobile").GetROProperty("checked") =1 Then
			If Browser("ContactInfoForBank_Step4").Page("ContactInfoForBank_Step4").WebCheckBox("Both").GetROProperty("checked") =1 Then
				logPass "Auto Check of BOTH - Both checked box has been auto checked by checking phone reminder & txtMessage options"
			Else
				logFail "Auto Check of BOTH - Both checked box has not been auto checked by checking phone reminder & txtMessage options"
			End If
		End If
		If  dictInData.Item("DT_AccType") = ACCTYPE_SBB Then
			Browser("ContactInfoForBank_Step4").Page("ContactInfoForBank_Step4").WebEdit("BusinessName").Set dictInData("DT_BusinessName")
		End If		
		If Browser("ContactInfoForBank_Step4").Page("ContactInfoForBank_Step4").WebElement("PhoneMeeting").Exist(2)Then
			Browser("ContactInfoForBank_Step4").Page("ContactInfoForBank_Step4").WebEdit("partyid").Set dictInData("DT_PartyID")
			Browser("ContactInfoForBank_Step4").Page("ContactInfoForBank_Step4").WebList("partyplatform").Select dictInData("DT_PlatForm")
		End If
	Else
		logFatal "Provide Contact Info for Bank - Page not Found"
	End If
End Function

'@@ 
'@ Name: submitContactInformation
'@ Description: Utility function to click on Submit\I Accept button in Provide Contact Information page
'@ Arg1: accType argument which contains the type of account user has selected
'Arg1: Possible account types are PB,SBB, IME or PPB
'@ Example: submitContactInformation("accType")
'@ Return:None
'@ History:
'@ Tags:
'@@
Public Function submitContactInformation(accType)
	If Browser("ContactInfoForBank_Step4").Page("ContactInfoForBank_Step4").Exist(1) Then
		'Click on Submit if banktype is not PPB
		If accType <> ACCTYPE_PPB Then
			Browser("ContactInfoForBank_Step4").Page("ContactInfoForBank_Step4").Link("Submit").Click
			'Click on Accept Button if PPB
		Else
			Browser("ContactInfoForBank_Step4").Page("ContactInfoForBank_Step4").Link("IAccept").Click
		End If
		CustomSync Browser("Appointment Confirmation Page").Page("Appointment Confirmation Page"), "False", "Submitted Contact Information"
	End If
End Function

'@@ 
'@ Name: cancelContactInformation
'@ Description: Utility function to click on Cancel button in Provide Contact Information page and asking user to continue or confirm cancellation
'@ Arg1: cancelOrContinue argument based on which application workflow either cancel or continues scheduling the appointment.
'@ Example: cancelContactInformation("cancel")
'@ Return:None
'@ History:
'@ Tags:
'@@
Public Function cancelContactInformation(cancelOrContinue)
	If Browser("ContactInfoForBank_Step4").Page("ContactInfoForBank_Step4").Exist(1) Then
		Browser("ContactInfoForBank_Step4").Page("ContactInfoForBank_Step4").Link("Cancel_ApptScheduling").Click
		Select Case cancel_or_continue
			Case CANCEL_APPOINTMENT
				Browser("ContactInfoForBank_Step4").Page("ContactInfoForBank_Step4").WebElement("CancelAppointment").Click
			Case CONTINUE_APPOINTMENT
				Browser("ContactInfoForBank_Step4").Page("ContactInfoForBank_Step4").WebElement("ContinueSchedulingAppointment").Click
		End Select		
	End If
End Function

'@@ 
'@ Name: checkAppointmentConfirmation
'@ Description: Utility function to check whether appointment confirmation page has been displayed or not after scheduling the appointment
'@ Example: checkAppointmentConfirmation("cancel")
'@ Return:None
'@ History:
'@ Tags:
'@@
Public Function checkAppointmentConfirmation()
	If Browser("Appointment Confirmation Page").Page("Appointment Confirmation Page").WebElement("ScheduleSuccess").Exist(10) Then
		'Browser("Appointment Confirmation Page").Page("Appointment Confirmation Page").WebTable("Confirmation number:").Highlight
		logPass "Appointment Confirmation Page - Appointment Confirmation Page displayed Successfully"
		Browser("Appointment Confirmation Page").Close
		CustomSync Browser("SmartLobby").Page("SmartLobby"), False, "Closed Confirmation Page"
		'To Do - Enhance Function :Add validations
	Else
		logFatal "Appointment Confirmation Page - Appointment Confirmation Page not displayed:"
	End If
End Function

'@@ 
'@ Name: getRecordinCurrentModule
'@ Description: To search for the record in the respective module so that the record gets available for display to get selected
'@ Return: None
'@ Example: getRecordinCurrentModule(DictionaryObject)
'@ History:
'@ Tags:
'@@
Public Function getRecordinCurrentModule(dictInData)
   Select Case CURRENT_MODULE
		Case MODULE_ACTIVITYMONITORING			
			'Default Search - Search with Date			
				If Browser("SmartLobby").Page("SmartLobby").Frame("Frame_ActivityManager").WebEdit("edDate_DD").Exist(1) Then						
					temp = Split(retrieveFromCache("DT_ApptDate"),"-")
					setMM = temp(0)
					setDD = temp(1)
					setYYYY = temp(2)
					If Len(setMM) =1 Then
						setMM = "0" & setMM
					End If
					If Len(setDD) = 1 Then
						setDD = "0" & setDD
					End If
					'To Check
					Browser("SmartLobby").Page("SmartLobby").Frame("Frame_ActivityManager").WebEdit("edDate_DD").NativeSet setDD				
					Browser("SmartLobby").Page("SmartLobby").Frame("Frame_ActivityManager").WebEdit("edMonth_MM").NativeSet setMM				
					Browser("SmartLobby").Page("SmartLobby").Frame("Frame_ActivityManager").WebEdit("edYear_YYYY").NativeSet setYYYY
					Browser("SmartLobby").Page("SmartLobby").Frame("Frame_ActivityManager").WebButton("Search").Click
				End If
		Case MODULE_APPOINTMENTMANAGER
			Browser("SmartLobby").Page("SmartLobby").Image("AppointmentManager").Click
			expected_Customer = dictInData.Item("DT_LastName") & retrieveFromCache("DT_CountLastFirstName") & ", " &_
													dictInData.Item("DT_FirstName") & retrieveFromCache("DT_CountLastFirstName")
			Browser("SmartLobby").Page("SmartLobby").Frame("Frame_AppointmentManager").WebEdit("Customer").Set expected_Customer
			Select Case CURRENT_WORKFLOW
				Case WORKFLOW_SCHEDULE
					'Browser("Smart Lobby").Page("Smart Lobby").Image("Appointment Manager").Click
					Browser("SmartLobby").Page("SmartLobby").Frame("Frame_AppointmentManager").WebButton("Search").Click
					'Browser("Smart Lobby").Page("Smart Lobby").Frame("Frame_AppointmentManager").WbfGrid("ApptDataGrid").Init
				Case WORKFLOW_WALKIN					
					Browser("SmartLobby").Page("SmartLobby").Frame("Frame_AppointmentManager").WebList("status").Select("Walk-in")
					Browser("SmartLobby").Page("SmartLobby").Frame("Frame_AppointmentManager").WebButton("Search").Click
					'Browser("Smart Lobby").Page("Smart Lobby").Frame("Frame_AppointmentManager").WbfGrid("ApptDataGrid").Init
				End Select
				Browser("SmartLobby").Page("SmartLobby").Frame("Frame_AppointmentManager").WbfGrid("ApptDataGrid").Init
   End Select
End Function

'@@ 
'@ Name: selectCreatedCustomerRecord
'@ Description: Utility function to validate whether schedule appointment has created a record and selecting the record in Scheduler Application.
'@ Arg1: dictInData Dictionary object which contains user information like lastName,firstName,email etc
'@ Return: None
'@ Example: selectCreatedCustomerRecord(DictionaryObject)
'@ History:
'@ Tags:
'@@
Public Function selectCreatedCustomerRecord(dictInData)
	'Get Expected Data
	expected_Customer = dictInData.Item("DT_LastName") & retrieveFromCache("DT_CountLastFirstName") & ", " &_
													dictInData.Item("DT_FirstName") & retrieveFromCache("DT_CountLastFirstName")
	expected_Location = Browser("SmartLobby").Page("SmartLobby").Frame("Frame_GENERAL").WebElement("LocationAddress").GetROProperty("innertext")
	expected_ApptDate = retrieveFromCache("DT_ApptDate")
	'expected_ApptTime = Replace(DataTable.Value("DT_ApptTime","TESTDATA")," ","")
	'expected_Role = retrieveFromCache.Item("Role")
	isFound = False
	Select Case CURRENT_MODULE
		Case MODULE_ACTIVITYMONITORING
			Set DESC_OBJ = Browser("SmartLobby").Page("SmartLobby").Frame("Frame_ActivityManager").WebTable("webTable_ActivityMonitor")
			'Get Data from Application	
			rCount = DESC_OBJ.RowCount
			cCount = DESC_OBJ.ColumnCount(1)
			'Get Customer Column Header Index
			For col=1 to cCount
				If UCASE("Customer") = UCASE(DESC_OBJ.GetCellData(1,col))Then
					customerIndex = col
					Exit For
				End If
			Next
			'Check for customer Data
			For i=2 to rCount
				If UCASE(DESC_OBJ.GetCellData(i,customerIndex)) = UCASE(expected_Customer) Then
					'Trim(expected_ApptTime) = Trim(DESC_OBJ.GetCellData(i,4)) AND _
					DESC_OBJ.Object.Rows(i-1).Click		
					addToCache "DT_RecordNumber", i
					logPass "RECORD VERIFCIATION - Record has been created successfully"
					isFound = True
					Exit For
				End If
			Next
			Case MODULE_APPOINTMENTMANAGER
				Set DESC_OBJ = Browser("SmartLobby").Page("SmartLobby").Frame("Frame_AppointmentManager").WbfGrid("ApptDataGrid")
				'Get Data from Application				
				rCount = DESC_OBJ.RowCount
				cCount = DESC_OBJ.ColumnCount(1)
				For i=2 to rCount
					If UCASE(DESC_OBJ.GetCellData(i,1)) = UCASE(expected_Customer) AND _
						Instr(1,expected_Location,DESC_OBJ.GetCellData(i,2)) >0 AND _
						UCase(DESC_OBJ.GetCellData(i,9)) = "SCHD" AND _
						FormatDateTime(expected_ApptDate) = FormatDateTime(DESC_OBJ.GetCellData(i,3)) Then
						'Trim(expected_Role) = Trim(DESC_OBJ.GetCellData(i,5)) 
						'Trim(expected_ApptTime) = Trim(DESC_OBJ.GetCellData(i,4)) AND _
						DESC_OBJ.Object.Rows(i-1).Click						
						addToCache "DT_RecordNumber", i
						logPass "RECORD VERIFCIATION - Record has been created successfully"
						isFound = True
						Exit For
					End If
				Next
	End Select
	If isFound Then
		logPass "Newly created customer " & expected_Customer & " found in appointment grid."
	Else
		logFatal "Newly created customer " & expected_Customer & " was not found in appointment grid."
	End If
End Function

'@@ 
'@ Name: goToReScheduleAppointment
'@ Description: Utility function to click on Re-Schedule button by checking whether Reschedule button is enabled or not
'@ Return: None
'@ Example: goToReScheduleAppointment
'@ History:
'@ Tags:
'@@
Public Function goToReScheduleAppointment()
   accType = retrieveFromDiscussionTopic("DT_AccType")
	Select Case CURRENT_MODULE
		Case MODULE_APPOINTMENTMANAGER
			If Browser("SmartLobby").Page("SmartLobby").Frame("Frame_AppointmentManager").Exist(1) Then
				isDisabled = Browser("SmartLobby").Page("SmartLobby").Frame("Frame_AppointmentManager").Link("Reschedule").GetROProperty("isDisabled")
				If isDisabled = True Then
					logFatal "Reschedule Button- You have not selected the record for rescheduling"
				Else
					Browser("SmartLobby").Page("SmartLobby").Frame("Frame_AppointmentManager").Link("Reschedule").Click
					If accType<> "PPB" Then
						CustomSync Browser("SelectDateAndTime_Step3").Page("SelectDateAndTime_Step3"),False,"Launched SelectDateTime Page"
					Else
						CustomSync Browser("SelectBankTopic_Step2").Page("SelectBankTopic_Step2"), False, "Launched Select Topic for Discussion page"
						Browser("SelectBankTopic_Step2").Page("SelectBankTopic_Step2").Link("Continue").Click
						CustomSync Browser("SelectDateAndTime_Step3").Page("SelectDateAndTime_Step3"),False,"Launched SelectDateTime Page"
					End If
				End If
			End If
		Case MODULE_ACTIVITYMONITORING
			If Browser("SmartLobby").Page("SmartLobby").Frame("Frame_ActivityManager").Exist(1) Then
				If Browser("SmartLobby").Page("SmartLobby").Frame("Frame_ActivityManager").WebElement("Reschedule").Exist(1) Then
					isDisabled = Browser("SmartLobby").Page("SmartLobby").Frame("Frame_ActivityManager").WebElement("Reschedule").GetROProperty("isDisabled")
					If isDisabled = False Then
						Browser("SmartLobby").Page("SmartLobby").Frame("Frame_ActivityManager").WebElement("Reschedule").Click
						If accType<> "PPB" Then
							CustomSync Browser("SelectDateAndTime_Step3").Page("SelectDateAndTime_Step3"),False,"Launched SelectDateTime Page"
						Else
							CustomSync Browser("SelectBankTopic_Step2").Page("SelectBankTopic_Step2"), False, "Launched Select Topic for Discussion page"
							Browser("SelectBankTopic_Step2").Page("SelectBankTopic_Step2").Link("Continue").Click
							CustomSync Browser("SelectDateAndTime_Step3").Page("SelectDateAndTime_Step3"),False,"Launched SelectDateTime Page"
						End If
					Else
						logFatal "Reschedule Button is Diasbled", "You have not selected the record for rescheduling"
					End If
				Else
					logFatal "ReSchedule Appointment - Button not available in Activity Monitoring tab.Please recheck"
				End If
			End If
		Case Else
			logFatal CURRENT_MODULE & "not avalable in application"
	End Select
	If Browser("SmartLobby").Dialog("Message from webpage").Exist(1) Then
		ERR_MSG = Browser("SmartLobby").Dialog("Message from webpage").Static.GetRoProperty("text")
		logFail ERR_MSG
	End If
	If  Browser("SelectDateAndTime_Step3").Page("SelectDateAndTime_Step3").Exist(1) Then
		logPass  "Reschedule Appointment - Launched Reschedule Appointment.Select Date and Time for Rescheduling"
	ElseIf Browser("Bank by Appointment").Page("Bank by Appointment").WebElement("There was a problem processing").Exist(1)Then
		ERR_MSG = Browser("Bank by Appointment").Page("Bank by Appointment").WebElement("There was a problem processing").GetROProperty("text")
		logFail ERR_MSG
	End If	
End Function

'@@ 
'@ Name: goToWalkIn
'@ Description: Utility function to click on Walkin button and enter user firstname,lastname in Select Discussion Topic page for walkin workflow.
'@Arg1 : dictInData which contains user firstname & lastname to be selected for workflow.
'@ Return: None
'@ Example: goToWalkIn(dictionaryObject)
'@ History:
'@ Tags:
'@@
Public Function goToWalkIn(dictInData)
    CountLastFirstName = RandomNumber(1,100)
	addToCache "DT_CountLastFirstName",CountLastFirstName
	If Browser("SmartLobby").Page("SmartLobby").Frame("Frame_ActivityManager").WebElement("Walk-In").Exist(1) Then
		isDisabled =  Browser("SmartLobby").Page("SmartLobby").Frame("Frame_ActivityManager").WebElement("Walk-In").Object.disabled
		If isDisabled=False Then
			Browser("SmartLobby").Page("SmartLobby").Frame("Frame_ActivityManager").WebElement("Walk-In").Click
			CustomSync Browser("SelectBankTopic_Step2").Page("SelectBankTopic_Step2"), False, "Launched Discussion Topic Page" 
			'Set current workflow as Walkin
			setCurrentWorkFlow WORKFLOW_WALKIN
			'Set App Date as Today Date
			walkinDate = Date
			walkinDate = Replace(walkinDate,"/","-")
			addToCache "DT_ApptDate", walkinDate			
		Else
			LogFatal "Object is Disabled - Fail"
		End If
		If Browser("SelectBankTopic_Step2").Page("SelectBankTopic_Step2").Exist(10) Then
			Browser("SelectBankTopic_Step2").Page("SelectBankTopic_Step2").WebEdit("FirstName").Set dictInData.Item("DT_FirstName") & CountLastFirstName 
			Browser("SelectBankTopic_Step2").Page("SelectBankTopic_Step2").WebEdit("LastName").Set dictInData.Item("DT_LastName") & CountLastFirstName
		End If
	Else
		logFatal "WALK-IN - Walkin Button not available"
	End If
End Function

'@@ 
'@ Name: handOff
'@ Description: Utility function to click on HandOff button after validating whether it is enabled or not
'@ Return: None
'@ Example: handOff
'@ History:
'@ Tags:
'@@
Public Function handOff()
	If Browser("SmartLobby").Page("SmartLobby").Frame("Frame_ActivityManager").Exist(1) Then		
		isDisabled = Browser("SmartLobby").Page("SmartLobby").Frame("Frame_ActivityManager").WebElement("Handoff").Object.disabled
		If isDisabled = False Then
			Browser("SmartLobby").Page("SmartLobby").Frame("Frame_ActivityManager").WebElement("Handoff").Click
			logPass "HandOff button has been clicked"
			CustomSync Browser("SelectBankTopic_Step2").Page("SelectBankTopic_Step2"), "False", "SelectTopic For Discussion page has been displayed"			
			Browser("SelectBankTopic_Step2").Page("SelectBankTopic_Step2").WebElement("Submit").Click
			CustomSync Browser("PostVisitSummary").Page("PostVisitSummary"), False, "Post Visit Summary page has been displayed"
		Else
			logFatal "HandOff button is disabled.Either record has not selected or selected record appointment is not of walkin workflow"
		End If
	End If
	'Post Summary Visit page
	If Browser("PostVisitSummary").Page("PostVisitSummary").Exist(1) Then
		Browser("PostVisitSummary").Page("PostVisitSummary").WebCheckBox("CommitCaseSubmitted").Set "ON" 'CheckBox DataTable.Value("PostSummary_CommitCase","TESTDATA")
		Browser("PostVisitSummary").Page("PostVisitSummary").WebButton("OK").Click
		CustomSync Browser("SmartLobby").Page("SmartLobby"), False, "Post Summary Visit page closed"
		'ToDo - Check manual work Flow
	Else
		logFail "POST VISIT SUMMARY - Post Visit sumamry page has not been displayed"
	End If
	'Check HandOff has been successful or not
	rValue = CInt(retrieveFromCache("DT_RecordNumber"))
	colCount = Browser("SmartLobby").Page("SmartLobby").Frame("Frame_ActivityManager").WebTable("webTable_ActivityMonitor").ColumnCount(1)
	For j=1 to colCount
		If UCASE("Call To Office") = UCASE(Browser("SmartLobby").Page("SmartLobby").Frame("Frame_ActivityManager").WebTable("webTable_ActivityMonitor").GetCellData(1,j)) Then
			colIndex = j
			Exit For
		End If
	Next
	time_CallToOffice = Browser("SmartLobby").Page("SmartLobby").Frame("Frame_ActivityManager").WebTable("webTable_ActivityMonitor").GetCellData(rValue,colIndex)
	If time_CallToOffice = " " OR time_CallToOffice = Empty Then
		logPass  "HandOff has been done successfully"
	Else
		logFail "HandOff failed.Please check the application workflow"
	End If
End Function

'@@ 
'@ Name: apptCheckIn
'@ Description: Utility function to click on ApptCheckIn button after validating whether it is enabled or not
'@ Return: None
'@ Example: apptCheckIn
'@ History:
'@ Tags:
'@@
Public Function apptCheckIn()
	If Browser("SmartLobby").Page("SmartLobby").Frame("Frame_ActivityManager").Exist(1) Then
		isDisabled = Browser("SmartLobby").Page("SmartLobby").Frame("Frame_ActivityManager").WebElement("Appt Check-In").Object.disabled
		If isDisabled = False Then
			Browser("SmartLobby").Page("SmartLobby").Frame("Frame_ActivityManager").WebElement("Appt Check-In").Click
			logPass "Clicked on Appointment checkin button"
		Else
			LogFatal "Appointment check In button is disabled.Either record has not selected or selected record appointment is of further dates"			
		End If
	End If
	'Check Appt has been checked in or not
	rValue = CInt(retrieveFromCache("DT_RecordNumber"))
	colCount = Browser("SmartLobby").Page("SmartLobby").Frame("Frame_ActivityManager").WebTable("webTable_ActivityMonitor").ColumnCount(1)
	'Traverse to the columns to get the column index of header
	For j=1 to colCount
		If UCASE("Check-In") = UCASE(Browser("SmartLobby").Page("SmartLobby").Frame("Frame_ActivityManager").WebTable("webTable_ActivityMonitor").GetCellData(1,j)) Then
			colIndex = j
			Exit For
		End If
	Next
	apptCheckin = Browser("SmartLobby").Page("SmartLobby").Frame("Frame_ActivityManager").WebTable("webTable_ActivityMonitor").GetCellData(rValue,colIndex)
	If apptCheckin <> " " Then
		logPass "Appointment has been checked in successfully"
	Else
		logFail "Appointment Check-In has been failed"
	End If
End Function

'@@ 
'@ Name: callToOffice
'@ Description: Utility function to click on callToOffice button after validating whether it is enabled or not
'@ Return: None
'@ Example: callToOffice
'@ History:
'@ Tags:
'@@
Public Function callToOffice()
	If Browser("SmartLobby").Page("SmartLobby").Frame("Frame_ActivityManager").Exist(1) Then		
		isDisabled = Browser("SmartLobby").Page("SmartLobby").Frame("Frame_ActivityManager").WebElement("Call to Office").Object.disabled
		If isDisabled = False Then
			Browser("SmartLobby").Page("SmartLobby").Frame("Frame_ActivityManager").WebElement("Call to Office").Click
			logPass "Call to Office button has been clicked"
		Else
			logFatal "Call to Office button is disabled.Either record has not selected or selected record appointment is of further dates"
		End If
	End If
	'Check Call to Office has been made or not	
	rValue = CInt(retrieveFromCache("DT_RecordNumber"))
	colCount = Browser("SmartLobby").Page("SmartLobby").Frame("Frame_ActivityManager").WebTable("webTable_ActivityMonitor").ColumnCount(1)
	For j=1 to colCount
		If UCASE("Call To Office") = UCASE(Browser("SmartLobby").Page("SmartLobby").Frame("Frame_ActivityManager").WebTable("webTable_ActivityMonitor").GetCellData(1,j)) Then
			colIndex = j
			Exit For
		End If
	Next
	time_CallToOffice = Browser("SmartLobby").Page("SmartLobby").Frame("Frame_ActivityManager").WebTable("webTable_ActivityMonitor").GetCellData(rValue,colIndex)
	If isEmpty(time_CallToOffice) = False Then
		logPass  "CallToOffice field is popultaed with current time"
	Else
		logFail "CallToOffice field value is not Empty"
	End If
End Function

'@@ 
'@ Name: serviceComplete
'@ Description: Utility function to click on serviceComplete button after validating whether it is enabled or not
'@ Return: None
'@ Example: serviceComplete
'@ History:
'@ Tags:
'@@
Public Function serviceComplete()
	If Browser("SmartLobby").Page("SmartLobby").Frame("Frame_ActivityManager").Exist(1) Then

		isDisabled = Browser("SmartLobby").Page("SmartLobby").Frame("Frame_ActivityManager").WebElement("Service Complete").Object.disabled
		If isDisabled = False Then
			Browser("SmartLobby").Page("SmartLobby").Frame("Frame_ActivityManager").WebElement("Service Complete").Click
			CustomSync Browser("PostVisitSummary").Page("PostVisitSummary"), False, "Clicked on Service Complete.Launched Post Summary Visit page"
			logPass  "Clicked on Service Complete button"
		Else
			logFatal "Service Complete Button is in disabled state.Either record has not selected or selected record appointment is of further dates"
		End If
	End If
	Select Case CURRENT_WORKFLOW
		Case WORKFLOW_SCHEDULE
			'Check Post Visit Summary
			If Browser("PostVisitSummary").Page("PostVisitSummary").Exist(1) Then
				Browser("PostVisitSummary").Page("PostVisitSummary").WebCheckBox("CommitCaseSubmitted").Set "ON" 
				Browser("PostVisitSummary").Page("PostVisitSummary").WebCheckBox("FollowUpApptMade").Set "ON" 
				Browser("PostVisitSummary").Page("PostVisitSummary").WebCheckBox("MeetTheCustomer").Set "OFF" 
				'Browser("PostVisitSummary").Page("PostVisitSummary").WebCheckBox("ReferralMade").set "OFF" 'UDF_SetCheckBox DataTable.Value("PostSummary_ReferalMade","TESTDATA")
				'Browser("PostVisitSummary").Page("PostVisitSummary").WebCheckBox("RelationShipReview").Set "OFF" 'UDF_SetCheckBox DataTable.Value("PostSummary_RelationShipReview","TESTDATA")
				'Browser("PostVisitSummary").Page("PostVisitSummary").WebCheckBox("SaleProcessed").set "ON" 'UDF_SetCheckBox DataTable.Value("PostSummary_SaleProcessed","TESTDATA")
				Browser("PostVisitSummary").Page("PostVisitSummary").WebButton("OK").Click
				CustomSync Browser("SmartLobby").Page("SmartLobby"), False, "Post summary visit completed"
				'ToDo - Check manual work Flow
			Else
				logFail "POST VISIT SUMMARY - Post Visit sumamry page has not been displayed"
			End If
		Case WORKFLOW_WALKIN
			'Check if required
	End Select
End Function


Public Function goToAssignAssociate()
	Select Case CURRENT_MODULE
		Case MODULE_ACTIVITYMONITORING
			If Browser("SmartLobby").Page("SmartLobby").Frame("Frame_ActivityManager").Exist(1) Then
				isDisabled = Browser("SmartLobby").Page("SmartLobby").Frame("Frame_ActivityManager").WebElement("Assign Associate").Object.disabled
				If isDisabled = False Then
					Browser("SmartLobby").Page("SmartLobby").Frame("Frame_ActivityManager").WebElement("Assign Associate").Click			
					logPass "Assign Associate button has been clicked"
					customSync Browser("Assign Associate").Page("Assign Associate"), False, "Assign Associate page has been launched"
				Else
					logFatal  "Assign Associate button is disabled.Either record has not selected or selected record appointment is of further dates"
				End If
			End If
		Case MODULE_APPOINTMENTMANAGER
			If Browser("SmartLobby").Page("SmartLobby").Frame("Frame_AppointmentManager").Exist(1) Then
				isDisabled = Browser("SmartLobby").Page("SmartLobby").Frame("Frame_AppointmentManager").Link("Assign Associate").Object.disabled
				If isDisabled = False Then
					Browser("SmartLobby").Page("SmartLobby").Frame("Frame_AppointmentManager").Link("Assign Associate").Click
					logPass "Assign Associate button has been clicked"
					customSync Browser("Assign Associate").Page("Assign Associate"), False, "Assign Associate page has been launched"
				Else
					logFatal  "Assign Associate button is disabled.Either record has not selected or selected record appointment is of further dates"
				End If
			End If
	End Select
End Function

Public Function assignAssociate(associateName)
	If Browser("Assign Associate").Page("Assign Associate").Exist(1) Then
		If Ucase(associateName) = UCase("selectAny") Then
			associatesList = Browser("Assign Associate").Page("Assign Associate").WebList("SelectAssociate").GetROProperty("all items")
			associatesList = Split(associatesList,";")
			If UBound(associatesList) >0 Then
				temp = RandomNumber(1,UBound(associatesList))
				wait(1)
				Browser("Assign Associate").Page("Assign Associate").WebList("SelectAssociate").Select associatesList(temp)
				addToCache "DT_AssociateName", associatesList(temp)
				Browser("Assign Associate").Page("Assign Associate").WebButton("Save Changes").Click
			Else
				logWarning "No associates in the list to select"
			End If
		Else
			Browser("Assign Associate").Page("Assign Associate").WebList("SelectAssociate").Select(associateName)			
			addToCache "DT_AssociateName", associateName
			Browser("Assign Associate").Page("Assign Associate").WebButton("Save Changes").Click
		End If
		CustomSync Browser("SmartLobby").Page("SmartLobby"), "False", "Changes saved successfully in Assign Associate page"	
	Else
		logFatal "Assign Associate page has not been dispalyed"
	End If
End Function

Public Function verifyAssociateName()
	rValue = CInt(retrieveFromCache("DT_RecordNumber"))
	expectedAssociateName = retrieveFromCache("DT_AssociateName")
	Select Case CURRENT_MODULE
		Case MODULE_ACTIVITYMONITORING
			If Browser("SmartLobby").Page("SmartLobby").Frame("Frame_ActivityManager").Exist(1) Then
				colCount = Browser("SmartLobby").Page("SmartLobby").Frame("Frame_ActivityManager").WebTable("webTable_ActivityMonitor").ColumnCount(1)
				For j=1 to colCount
					If UCase("Associate") = UCase(Browser("SmartLobby").Page("SmartLobby").Frame("Frame_ActivityManager").WebTable("webTable_ActivityMonitor").GetCellData(1,j)) Then
						colIndex = j
						Exit For
					End If
				Next
				'Browser("Smart Lobby").Page("Smart Lobby").Frame("Frame_ActivityManager").WebTable("webTable_ActivityMonitor").Object.rows(rValue-1).Click
				actualAssociateName = Browser("SmartLobby").Page("SmartLobby").Frame("Frame_ActivityManager").WebTable("webTable_ActivityMonitor").GetCellData(rValue,colIndex)
			End If
		Case MODULE_APPOINTMENTMANAGER
			If Browser("SmartLobby").Page("SmartLobby").Frame("Frame_AppointmentManager").Exist(1) Then
				colCount = Browser("SmartLobby").Page("SmartLobby").Frame("Frame_AppointmentManager").WbfGrid("ApptDataGrid").ColumnCount(1)
				For j=1 to colCount
					If UCase("Associate") = UCase(Browser("SmartLobby").Page("SmartLobby").Frame("Frame_AppointmentManager").WbfGrid("ApptDataGrid").GetCellData(1,j)) Then
						colIndex = j
						Exit For
					End If
				Next
				Browser("SmartLobby").Page("SmartLobby").Frame("Frame_AppointmentManager").WbfGrid("ApptDataGrid").Object.rows(rValue-1).Click
				actualAssociateName = Browser("SmartLobby").Page("SmartLobby").Frame("Frame_AppointmentManager").WbfGrid("ApptDataGrid").GetCellData(rValue,colIndex)
			End If
	End Select
	If UCase(actualAssociateName) = UCase(expectedAssociateName) Then
		logPass  "Assign Associate - Actual and expected associated names matching."
	Else
		logFail "Assign Associate - Actual and expected associated names not  matching."
	End If
End Function

'@@ 
'@ Name: checkRecordInCompleteTab
'@ Description: Utility function to create a followup appointment based on user input
'@ Arg1:After service complete to check the record in Complete tab in Activity Monitoring module.
'@ Return: None
'@ Example: checkRecordInCompleteTab(dictionayObject)
'@ History:
'@ Tags:
'@@
Public Function checkRecordInCompleteTab(dictInData)
	'Check the record in COMPLETE Tab
	If Browser("SmartLobby").Page("SmartLobby").Frame("Frame_ActivityManager").Exist(1) Then
		Browser("SmartLobby").Page("SmartLobby").Frame("Frame_ActivityManager").Link("Tab_Complete").Click		
		'Get Expected Data
		expected_Customer = dictInData.Item("DT_LastName") & retrieveFromCache("DT_CountLastFirstName") & ", " &_	
														dictInData.Item("DT_FirstName") & retrieveFromCache("DT_CountLastFirstName")
		'expected_Role = DataTable.Value("Role","TESTDATA")
		isFound = False
		set DESC_OBJ = Browser("SmartLobby").Page("SmartLobby").Frame("Frame_ActivityManager").WebTable("webTable_ActivityMonitor")
		rCount = DESC_OBJ.RowCount
		For i=2 to rCount
			If UCASE(DESC_OBJ.GetCellData(i,4)) = UCASE(expected_Customer) Then 
				'Trim(expected_Role) = Trim(DESC_OBJ.GetCellData(i,3)) Then
				'Trim(expected_ApptTime) = Trim(DESC_OBJ.GetCellData(i,4)) AND _
				Browser("SmartLobby").Page("SmartLobby").Frame("Frame_ActivityManager").WebTable("webTable_ActivityMonitor").Object.Rows(i-1).Click
				logPass "RECORD VERIFCIATION - Record has been moved successfully to Complete Tab"
				isFound = True
				Exit For
			End If
		Next
		If isFound = False Then
			logFail "Record not found in Complete tab"		
		End If
	Else
		logFatal "Not in Activity Monitoring Module to select Complete tab"
	End If
End Function

'@@ 
'@ Name: createFollowUpAppointment
'@ Description: Utility function to create a followup appointment based on user input
'@ Arg1:appointment type to be selected -whetehr new or Followup type
'@ Arg 2: SelfAssign "Yes" or "No" based on which selfAssign radio button will be set
'@ Return: None
'@ Example: createFollowUpAppointment(APPOINTMENT_NEW,"No")
'@ History:
'@ Tags:
'@@

Public Function createFollowUpAppointment() ',selfAssign)
   Browser("SmartLobby").Page("SmartLobby").Frame("Frame_ActivityManager").WebElement("Schedule").Click
   CustomSync Browser("Schedule_New_FollowUp_App").Page("Schedule_New_FollowUp_App"),False, "Launched Followup Appointment page"
	If Browser("Schedule_New_FollowUp_App").Page("Schedule_New_FollowUp_App").Exist(1)  Then
		Select Case FOLLOWUPAPPT_TYPE
			Case APPOINTMENT_NEW
				Browser("Schedule_New_FollowUp_App").Page("Schedule_New_FollowUp_App").WebCheckBox("FollowUp").Set "OFF"
				Browser("Schedule_New_FollowUp_App").Page("Schedule_New_FollowUp_App").WebCheckBox("NewAppointment").Set "ON"					
			'Case APPOINTMENT_TYPE_FOLLOWUP
		End Select
		'Check Self Assign option and set value
		'If  selfAssign = "No" Then
		'	Browser("Schedule_New_FollowUp_App").Page("Schedule_New_FollowUp_App").WebRadioGroup("selfassign").Select "0"
		'End If
		'Click on Schedule Appointment button
		Browser("Schedule_New_FollowUp_App").Page("Schedule_New_FollowUp_App").WebButton("ScheduleAppointment").Click
		customSync Browser("SelectBankTopic_Step2").Page("SelectBankTopic_Step2"),False,"Launched Select Discussion Topic Page"
		logPass "FollowUp appointment can be created"
	End If
End Function

'@@ 
'@ Name: Details
'@ Description: Utility function to click on Details tab by validating whether tab is enabled or not
'@ Return: None
'@ Example: Details()
'@ History:
'@ Tags:
'@@
'Functions on Details page
Public Function Details()
	Select Case CURRENT_MODULE
		Case MODULE_ACTIVITYMONITORING
			If Browser("SmartLobby").Page("SmartLobby").Frame("Frame_ActivityManager").Exist(1) Then
				isDisabled = Browser("SmartLobby").Page("SmartLobby").Frame("Frame_ActivityManager").WebElement("Details").Object.isDisabled			
				If isDisabled = False Then
					Browser("SmartLobby").Page("SmartLobby").Frame("Frame_ActivityManager").WebElement("Details").Click
					'logPass "Clicked on Details Tab"
					CustomSync Browser("AppointmentDetails").Page("AppointmentDetails"), False, "Deatils page has been launched"
					logPass "Details button has been clicked and Details page has been launched"
				Else
					logFail "Details button is disabled.Please recheck application workflow"
				End If
			Else
				logFatal "Activity Monitor module is not available.Please recheck app workflow"
        	End If
		Case MODULE_APPOINTMENTMANAGER
			If Browser("SmartLobby").Page("SmartLobby").Frame("Frame_AppointmentManager").Exist(1) Then
				isDisabled = Browser("SmartLobby").Page("SmartLobby").Frame("Frame_AppointmentManager").Link("Details").Object.disabled
				If isDisabled = False Then
					Browser("SmartLobby").Page("SmartLobby").Frame("Frame_AppointmentManager").Link("Details").Click
					logPass "Clicked on Details Tab"
					CustomSync Browser("AppointmentDetails").Page("AppointmentDetails") , False, "Details page has been launched"					
				Else
					logFail "Details button is disabled.Please recheck application workflow"
           		End If
			Else
				logFatal "Appointment Manager module is not available.Please recheck app workflow"
			End If
	End Select
End Function

'@@ 
'@ Name: checkTabOptionsinDetailsPage
'@ Description: Utility function to check tab options(Details and Demographics) based on loginView
'@ Return: None
'@ Example: checkTabOptionsinDetailsPage()
'@ History:
'@ Tags:
'@@
Public Function checkTabOptionsinDetailsPage()
	If Browser("AppointmentDetails").Page("AppointmentDetails").Exist(1) Then
		Select Case CURRENT_VIEW
			Case VIEW_PLATFORM
				If Browser("AppointmentDetails").Page("AppointmentDetails").Link("Tab_Demographics").Exist(1) AND _
					Browser("AppointmentDetails").Page("AppointmentDetails").Link("Tab_Details").Exist(1) Then
					logPass "DETIALS -TAB OPTIONS - Both Demographics and Details tab are displaying in PlatForm View"
				Else
					logFail  "DETIALS -TAB OPTIONS - Tab options are not getting displayed correctly for PlatForm View"
				End If
			Case VIEW_LOBBY
				If  Browser("AppointmentDetails").Page("AppointmentDetails").Link("Tab_Details").Exist(1) and _
					Not(Browser("AppointmentDetails").Page("AppointmentDetails").Link("Tab_Demographics").Exist(1)) Then
						logPass "DETIALS -TAB OPTIONS - Only Details tab is getting displayed in Lobby View"					
				Else
					logFail  "DETIALS -TAB OPTIONS - Tab options are not getting displayed correctly for Lobby View"
				End If
		End Select
		Browser("AppointmentDetails").Page("AppointmentDetails").WebButton("Close").Click
		CustomSync Browser("SmartLobby").Page("SmartLobby"), False, "Closed Details page"
	End If
End Function

'@@ 
'@ Name: verifyDataInDetailsTab
'@ Description: Utility function to verify data in Details tab under Details page
'@ Arg1: Dictionary Object which contains firstName,lastName for verification
'@ Return: None
'@ Example: verifyDataInDetailsTab(dictionaryObject)
'@ History:
'@ Tags:
'@@
Public Function verifyDataInDetailsTab(dictInData)
	If Browser("AppointmentDetails").Page("AppointmentDetails").Exist(1) Then
		Browser("AppointmentDetails").Page("AppointmentDetails").Link("Tab_Details").Click
		If Browser("AppointmentDetails").Page("AppointmentDetails").WebElement("Details_Name").Exist(1) AND _
			Browser("AppointmentDetails").Page("AppointmentDetails").WebElement("Details_Staff DetailsHeader").Exist(1) Then
			logPass "GROUP OPTIONS in DETAILS tab - Details and Staff Details groups are available"
		Else
			logFail  "GROUP OPTIONS in DETAILS tab - Details and Staff Deatils groups are not available"
		End If
		'UserName Validation
		tempName = Browser("AppointmentDetails").Page("AppointmentDetails").WebElement("Details_Name").GetROProperty("innerhtml")
		expName = dictInData.Item("DT_LastName") & retrieveFromCache("DT_CountLastFirstName") & ", " &_	
														dictInData.Item("DT_FirstName") & retrieveFromCache("DT_CountLastFirstName")
		If UCASE(tempName) = UCASE(expName) Then
			logPass "USERNAME Validation - UserName Validation passed"
		Else
			logFail "USERNAME Validation - UserName Validation Failed"
		End If
		'Check Date Field is non Editable
		isDisabled = Browser("AppointmentDetails").Page("AppointmentDetails").WebEdit("Details_Date").Object.disabled
		If isDisabled=True Then
			logPass "DATE FIELD - Date Field is non Editable"
		Else
			logFail "DATE FIELD - Date Field is Editable"
		End If	
		'Close Details Page
		Browser("AppointmentDetails").Page("AppointmentDetails").WebButton("Close").Click		
		CustomSync Browser("SmartLobby").Page("SmartLobby"), False, "Closed Details page"
	End If
End Function

'@@ 
'@ Name: editDataInDemographicsTab
'@ Description: Utility function to edit data in Demographics tab in Details page
'@ Arg1: Dictionary Object which contains firstName,lastName,contact Number etc to be updated
'@ Return: None
'@ Example: editDataInDemographicsTab(dictionaryObject)
'@ History:
'@ Tags:
'@@
Public Function editDataInDemographicsTab(dictInData)
	If Browser("AppointmentDetails").Page("AppointmentDetails").Exist(1) Then
		Browser("AppointmentDetails").Page("AppointmentDetails").Link("Tab_Demographics").Click			
		'Update LastName
		If isEmpty(dictInData.Item("DT_LastName")) = False Then
			Browser("AppointmentDetails").Page("AppointmentDetails").WebEdit("Demographics_LastName").Set dictInData.Item("DT_LastName")
			'dictContactInfo.Add "DT_LastName",dictInData.Item("DT_LastName")
			addToCache "DT_LastName",dictInData.Item("DT_LastName")
		End If
		'Update FirstName
		If isEmpty(dictInData.Item("DT_FirstName")) =False Then
			Browser("AppointmentDetails").Page("AppointmentDetails").WebEdit("Demographics_FirstName").Set dictInData.Item("DT_FirstName")
			addToCache "DT_FirstName",dictInData.Item("DT_FirstName")
		End If
		'Update Email ID
		If isEmpty(dictInData.Item("DT_Email")) = False Then
			Browser("AppointmentDetails").Page("AppointmentDetails").WebEdit("Demographics_Email").Set dictInData.Item("DT_Email")
			addToCache "DT_Email",dictInData.Item("DT_Email")
		End If
		'Update Contact Number			
		If isEmpty(dictInData.Item("DT_ContactNumber")) = False Then
			temp = dictInData.Item("DT_ContactNumber")
			expected_ContactNum = MID(dictInData.Item("DT_ContactNumber"),1,3) & "-" & MID(dictInData.Item("DT_ContactNumber"),4,3) _
							& "-" & MID(dictInData.Item("DT_ContactNumber"),7,4)
			Browser("AppointmentDetails").Page("AppointmentDetails").WebEdit("Demographics_ContactNumber").set expected_ContactNum
			addToCache "DT_ContactNumber",expected_ContactNum	
		End If
		'Update Existing Customer check Box
		isDisabled = Browser("AppointmentDetails").Page("AppointmentDetails").WebCheckBox("Demographics_ExistCustomer").Object.disabled
		If isDisabled = False Then
			If UCASE(dictInData.Item("DT_ExistingCustomer")) = "YES" Then			
				Browser("AppointmentDetails").Page("AppointmentDetails").WebCheckBox("Demographics_ExistCustomer").Set "ON"
				addToCache "DT_ExistingCustomer","Yes"
			ElseIf UCASE(dictInData.Item("DT_ExistingCustomer")) = "NO" Then
				Browser("AppointmentDetails").Page("AppointmentDetails").WebCheckBox("Demographics_ExistCustomer").Set "OFF"
				addToCache "DT_ExistingCustomer","No"
			End If
		End If
		Browser("AppointmentDetails").Page("AppointmentDetails").WebButton("Save").Click
		
		'Check if any Error
		If Dialog("text:=Message from webpage").Exist(1) Then
			Dialog("text:=Message from webpage").WinButton("text:=OK").Click
			logFatal "ERROR - Invalid test Data.Check test data sheet"
		End If
		CustomSync Browser("SmartLobby").Page("SmartLobby"), False, "Data in Demographics tab has been updated"		
	End If
End Function

'@@ 
'@ Name: verifyDatainDemographicsTab
'@ Description: Utility function to verify Demographics tab in Details page
'@ Arg1: Dictionary Object which contains firstName,lastName,contact Number etc for validation
'@ Return: None
'@ Example: verifyDatainDemographicsTab(dictionaryObject)
'@ History:
'@ Tags:
'@@
Public Function verifyDatainDemographicsTab(dictInData)
   If Browser("AppointmentDetails").Page("AppointmentDetails").Exist(1) Then
		Browser("AppointmentDetails").Page("AppointmentDetails").Link("Tab_Demographics").Click
		'Expected Values - New schedule Appointment
		expectedLastName = retrieveFromCache("DT_LastName")
		expectedFirstName =retrieveFromCache("DT_FirstName")
		expectedEmail = retrieveFromCache("DT_Email")
		expectedContactNum = retrieveFromCache("DT_ContactNumber")
		expectedExistingCustomer = retrieveFromCache("DT_ExistingCustomer")		

		'Check for balnks - if edit data has not been done pick actual data
		If  isEmpty(expectedLastName) Then
			expectedLastName = dictInData.Item("DT_LastName") & retrieveFromCache("DT_CountLastFirstName")
		End If
		If  isEmpty(expectedFirstName) Then
			expectedFirstName = dictInData.Item("DT_FirstName") & retrieveFromCache("DT_CountLastFirstName")
		End If
		If  isEmpty(expectedEmail) Then
			expectedEmail = dictInData.Item("DT_Email")
		End If
		If  isEmpty(expectedContactNum) Then
			expectedContactNum = dictInData.Item("DT_ContactNumber")
		End If
		If  isEmpty(expectedExistingCustomer) Then
			expectedExistingCustomer = dictInData.Item("DT_ExistingCustomer")
		End If

		'Actual Values
		actualLastName = Browser("AppointmentDetails").Page("AppointmentDetails").WebEdit("Demographics_LastName").GetROProperty("value")
		actualFirstName = Browser("AppointmentDetails").Page("AppointmentDetails").WebEdit("Demographics_FirstName").GetROProperty("value")
		actualEmail = Browser("AppointmentDetails").Page("AppointmentDetails").WebEdit("Demographics_Email").GetROProperty("value")
		actualContactNum = Browser("AppointmentDetails").Page("AppointmentDetails").WebEdit("Demographics_ContactNumber").GetROProperty("value")
		actualContactNum = Replace(actualContactNum,"-","")
		actualExistingCustomer = Browser("AppointmentDetails").Page("AppointmentDetails").WebCheckBox("Demographics_ExistCustomer").GetROProperty("checked")

		If  actualExistingCustomer = 1 Then
			actualExistingCustomer = "Yes"
		Else
			actualExistingCustomer = "No"
		End If
		'Check actual and expected Data
		If  UCASE(expectedLastName) = UCASE(actualLastName) And _
			UCASE(expectedFirstName) = UCASE(actualFirstName) And _
			UCASE(expectedEmail) = UCASE(actualEmail) And _
			UCASE(expectedContactNum) = UCASE(actualContactNum) And _
			UCASE(expectedExistingCustomer) = UCASE(actualExistingCustomer) Then
			logPass "DEMOGRAPHICS DATA CHECK - Data Validation passed"
		Else
			logFail "DEMOGRAPHICS DATA CHECK - Data not matching with actual Data.Validation Failed"
		End If
		'Close Appointment Details tab
		Browser("AppointmentDetails").Page("AppointmentDetails").WebButton("Close").Click
		CustomSync Browser("SmartLobby").Page("SmartLobby"), False, "Closed Details page"
	End If
End Function

'@@
'@ Name: cancelAppointment
'@ Description: Utility function to click on Calcel appointment in the respectively module to cancel an appointment
'@ Return: None
'@ Example: cancelAppointment()
'@ History:
'@ Tags:
'@@
Public Function cancelAppointment()
   Select Case CURRENT_MODULE
		Case MODULE_ACTIVITYMONITORING
			If Browser("SmartLobby").Page("SmartLobby").Frame("Frame_ActivityManager").Exist(1) Then
				isDisabled = Browser("SmartLobby").Page("SmartLobby").Frame("Frame_ActivityManager").WebElement("Cancel").Object.disabled
				If isDisabled = False Then
					Browser("SmartLobby").Page("SmartLobby").Frame("Frame_ActivityManager").WebElement("Cancel").Click
					logPass "Cancel button has been clicked"
				Else
					logFail "Cancel button is disabled.Either record has not selected or selected record appointment is of earlier dates"
				End If
			Else
				logFatal "Currently not in Activity Monitor module.Please check the workflow"
			End If
		Case MODULE_APPOINTMENTMANAGER
			If Browser("SmartLobby").Page("SmartLobby").Frame("Frame_AppointmentManager").Exist(1) Then
				isDisabled = Browser("SmartLobby").Page("SmartLobby").Frame("Frame_AppointmentManager").Link("Cancel").Object.disabled
				If isDisabled = False Then
					Browser("SmartLobby").Page("SmartLobby").Frame("Frame_AppointmentManager").Link("Cancel").Click
					logPass "Cancel button has been clicked"
				Else
					logFail "Cancel button is disabled.Either record has not selected or selected record appointment is of earlier dates"
				End If
			Else
				logFatal "Currently not in Appointment manger module.Please check the workflow"				
			End If
   End Select
	'Check for confirmation Dialog Box
	If Dialog("text:=Message from webpage").Exist(1) Then
		Dialog("text:=Message from webpage").WinButton("text:=OK").Click
		wait 1
	End If
	'If still Dialog Exist - Cancel Appointment has failed
	If Dialog("text:=Message from webpage").Exist(1) Then
		Dialog("text:=Message from webpage").WinButton("text:=OK").Click
		logFatal "It is too late to Cancel the appointment -Cancellation failed"		
	Else
		logPass "Appointment has been cancelled successfully"
	End If
End Function

'@@
'@ Name: verifyAppointmentCancellation
'@ Description: Utility function to click on Calcel appointment in the respectively module to cancel an appointment
'@ Arg1: Dictionary Object which contains firstName,lastName,contact Number etc for validation
'@ Return: None
'@ Example: verifyAppointmentCancellation(dictionayObject)
'@ History:
'@ Tags:
'@@


Public Function verifyAppointmentCancellation()
		expected_Customer = retrieveFromContactInfo("DT_LastName") & retrieveFromCache("DT_CountLastFirstName") & ", " &_	
													retrieveFromContactInfo("DT_FirstName") & retrieveFromCache("DT_CountLastFirstName")
		recordFound = False
	Select Case CURRENT_MODULE
		Case MODULE_ACTIVITYMONITORING
			Browser("SmartLobby").Page("SmartLobby").Frame("Frame_ActivityManager").Link("Tab_Sign-In").Click
			Set DESC_OBJ = Browser("SmartLobby").Page("SmartLobby").Frame("Frame_ActivityManager").WebTable("webTable_ActivityMonitor")
			'Get Data from Application
			rCount = DESC_OBJ.RowCount
			cCount = DESC_OBJ.ColumnCount(1)
			For col=1 to cCount
				If UCASE("Customer") = UCASE(DESC_OBJ.GetCellData(1,col))Then
					customerIndex = col
					Exit For
				End If
			Next
			For i=2 to rCount
				If UCASE(DESC_OBJ.GetCellData(i,customerIndex)) = UCASE(expected_Customer) Then					
					recordFound = True
					Exit For
				End If
			Next
			If recordFound = True Then				
				logFail "Cancel Appointment Failed"
			Else
				logPass "Cancel Appointment Successfully Done"
			End If
		Case MODULE_APPOINTMENTMANAGER
			Set DESC_OBJ = Browser("SmartLobby").Page("SmartLobby").Frame("Frame_AppointmentManager").WbfGrid("ApptDataGrid")
			'Get Data from Application			
			rCount = DESC_OBJ.RowCount
			cCount = DESC_OBJ.ColumnCount(1)
			For i=2 to rCount
				If UCASE(DESC_OBJ.GetCellData(i,1)) = UCASE(expected_Customer) Then
					status = DESC_OBJ.GetCellData(i,9)
					If status = "CANC" Then
						logPass "CancelAppointment has been done succssfully .Status of record changed to CANC"	
						recordFound = True					
					Else
						logFail "CancelAppointment failed .Status of record not changed to CANC"						
					End If
					Exit For
				End If
             Next
			 If recordFound =False Then
				 logFail "CancelAppointment Failed. Record not Found"
			 End If
	End Select
End Function

'@@
'@ Name: noShow
'@ Description: Utility function to click on Calcel appointment in the respectively module to cancel an appointment
'@ Arg1: Dictionary Object which contains firstName,lastName,contact Number etc for validation
'@ Return: None
'@ Example: noShow(dictionaryObject)
'@ History:
'@ Tags:
'@@
Public Function noShow(dictInData)
	If Browser("SmartLobby").Page("SmartLobby").Frame("Frame_ActivityManager").Exist(1) Then
		isDisabled = Browser("SmartLobby").Page("SmartLobby").Frame("Frame_ActivityManager").WebElement("No Show").Object.disabled
		If isDisabled = False Then
			Browser("SmartLobby").Page("SmartLobby").Frame("Frame_ActivityManager").WebElement("No Show").Click			
			logPass "NO SHOW button has been clicked"
		Else
			logFatal  "NO SHOW button is disabled.Either record has not selected or selected record appointment is of further dates"
        End If
	End If
    'Click on No-Show tab
	Browser("SmartLobby").Page("SmartLobby").Frame("Frame_ActivityManager").Link("Tab_No-Show").Click
	'Get Expected Data
	expected_Customer = dictInData.Item("DT_LastName") & retrieveFromCache("DT_CountLastFirstName") & ", " &_	
														dictInData.Item("DT_FirstName") & retrieveFromCache("DT_CountLastFirstName")	
	isFound = False
	set DESC_OBJ = Browser("SmartLobby").Page("SmartLobby").Frame("Frame_ActivityManager").WebTable("webTable_ActivityMonitor")
	rCount = DESC_OBJ.RowCount
	For col=1 to cCount
		If UCASE("Customer") = UCASE(DESC_OBJ.GetCellData(1,col))Then
			customerIndex = col
			Exit For
		End If
	Next        
	For i=2 to rCount
		If UCASE(DESC_OBJ.GetCellData(i,customerIndex)) = UCASE(expected_Customer) Then					
			Browser("SmartLobby").Page("SmartLobby").Frame("Frame_ActivityManager").WebTable("webTable_ActivityMonitor").Object.Rows(i-1).Click
			isFound = True
			logPass "Record moved to No-Show tab successfully"
			Exit For
		End If
	Next
	If isFound = False Then
		logFail "NO SHOW - Validation Failed"
	End If   
End Function

'@@ 
'@ Name: checkObjectIsDisabled
'@ Description: To check whether an object is Disabled or not
'@ Arg1: inObj argument is the object to be validated whether Disabled or not
'@ Return: True if Disabled and false if not
'@ Example: checkObjectIsDisabled(inObj)
'@ History:
'@ Tags:
'@@
Public Function checkObjectIsDisabled(inObj)   
   VerifyProperty_ inObj,"disabled",True, inObj.GetRoProperty("text") & "button is in disabled state"
End Function

'@@ 
'@ Name: checkObjectIsEnabled
'@ Description: To check whether an object is Enabled or not
'@ Arg1: inObj argument is the object to be validated whether Enabled or not
'@ Return: True if enabled and false if not
'@ Example: checkObjectIsEnabled(inObj)
'@ History:
'@ Tags:
'@@
Public Function checkObjectIsEnabled(inObj)   
   VerifyProperty_ inObj,"disabled",False,inObj.GetRoProperty("text") & "button is in enabled state"
End Function

'@@ 
'@ Name: checkIsExist
'@ Description: To check whether an object exists or not
'@ Arg1: inObj argument is the object to be validated for existence
'@ Return: None
'@ Example: checkIsExist(InObj)
'@ History:
'@ Tags:
'@@
Public Function checkIsExist(inObj)
   If inObj.Exist(1) Then
	   logPass inObj.GetRoProperty("text") & "Exists"
	Else
		logFail inObj.GetRoProperty("text") & "does not Exists"
   End If
End Function

'@@ 
'@ Name: checkIsNotExist
'@ Description: To check whether an object exists or not
'@ Arg1: inObj argument is the object to be validated for existence
'@ Return: None
'@ Example: checkIsExist(InObj)
'@ History:
'@ Tags:
'@@
Public Function checkIsNotExist(inObj)
   If inObj.Exist(1) Then
	   logFail inObj.GetRoProperty("text") & "Exists"
	Else
		logPass inObj.GetRoProperty("text") & "does not Exists"
   End If
End Function

'@@ 
'@ Name: searchRecordinActivityMonitoring
'@ Description: To search for the record based on name or date or on both in Activity Monitoring module
'@ Arg1: searchCriteria based on which search needs to be performed
'@ Return: None
'@ Example: searchRecordinActivityMonitoring("searchCriteria")
'@ History:
'@ Tags:
'@@
Public Function searchRecordinActivityMonitoring(searchCriteria)
	If searchCriteria = SEARCHCRITERIA_NAME OR searchCriteria = SEARCHCRITERIA_NAMEnDATE Then
		'Search with Name
		If Browser("SmartLobby").Page("SmartLobby").Frame("Frame_ActivityManager").WebEdit("Name").Exist(1) Then
			Browser("SmartLobby").Page("SmartLobby").Frame("Frame_ActivityManager").WebEdit("Name").Set _
			Dict_ContactInfo.Item("DT_LastName") & Dict_RunTimeCache.Item("DT_CountLastFirstName") & ", " &_
			Dict_ContactInfo.Item("DT_FirstName") & Dict_RunTimeCache.Item("DT_CountLastFirstName")
		End If
	ElseIf searchCriteria = SearchCriteria_Date OR searchCriteria = SEARCHCRITERIA_NAMEnDATE Then
	End If
	Browser("SmartLobby").Page("SmartLobby").Frame("Frame_ActivityManager").WebButton("Search").Click
	Browser("SmartLobby").Page("SmartLobby").Frame("Frame_ActivityManager").Init
End Function

Public Function cancelServiceCompletedRecord
   Browser("SmartLobby").Page("SmartLobby").Frame("Frame_AppointmentManager").Link("Cancel").Click
   If Dialog("text:=Message from webpage").Exist(1) Then
	   Set myObj = Description.Create
	   myObj("micclass").value = "Static"
	   set chdItems = Dialog("text:=Message from webpage").ChildObjects(myObj)
	   msg = chdItems(1).GetRoProperty("text")
	   If Instr(1,msg, "Cannot cancel") Then
		   logPass "Cancel a scheduled appointment" & msg
		   Dialog("text:=Message from webpage").WinButton("text:=OK").Click
	   ElseIf Instr(1,msg,"Are you sure") Then
			Dialog("text:=Message from webpage").WinButton("text:=Cancel").Click
			logFail "Allowing Cancellation of service completed record .Please check workflow"	   
	   End If
	Else
		logFail "No message has been displayed while cancelling the service completed appointment"
   End If
End Function

'Check cancellation of service completed appointment
Public Function CancelAfterServiceCompleteWorkflow()
	goToModule MODULE_APPOINTMENTMANAGER
	getRecordinCurrentModule DICTCONTACTINFO
	selectCreatedCustomerRecord DICTCONTACTINFO
	cancelServiceCompletedRecord
End Function

Public Function scheduleAppointmentWorkFlow()
	goToScheduleAppointment
	selectTopicforDiscussion dictSelectTopicForDiscussion
	selectDateTime zipCode, dateToSelect
	provideContactInformation dictContactInfo
	submitContactInformation accType
	checkAppointmentConfirmation
	getRecordinCurrentModule dictContactInfo
	selectCreatedCustomerRecord dictContactInfo
End Function

Public Function walkInWorkFlow()
	goToWalkIn dictContactInfo
	selectTopicforDiscussion dictSelectTopicForDiscussion
	getRecordinCurrentModule dictContactInfo
	selectCreatedCustomerRecord dictContactInfo
End Function

Public Function reScheduleAppointmentWorkFlow()
   scheduleAppointmentWorkFlow
   goToReScheduleAppointment
   'In case of PPB reschedule launches select topic for bank page
	selectDateTime zipCode, dateToSelect
	'provideContactInformation dictContactInfo
	submitContactInformation accType
	checkAppointmentConfirmation
	getRecordinCurrentModule dictContactInfo
	selectCreatedCustomerRecord dictContactInfo
End Function

Public Function serviceCompleteWorkFlow()
   	callToOffice
	serviceComplete
	checkRecordInCompleteTab dictContactInfo
End Function

Public Function followUpWorkFlow()
   	createFollowUpAppointment()
	If  FOLLOWUPAPPT_TYPE = APPOINTMENT_NEW Then
		selectTopicforDiscussion dictSelectTopicForDiscussion
	Else
		Browser("SelectBankTopic_Step2").Page("SelectBankTopic_Step2").Link("Continue").Click
		CustomSync Browser("SelectDateAndTime_Step3").Page("SelectDateAndTime_Step3"), False, "Launched Select Date Time page"
	End If	
	'Get Actual Data
	If followUpScenario = "FollowUpAfterCallToOffice" Then
		tempCountLastFirstName = DICTRUNTIMECACHE.Item("DT_CountLastFirstName")
		tempLastName = DICTCONTACTINFO.Item("DT_LastName")
		tempFirstName = DICTCONTACTINFO.Item("DT_FirstName")
		tempFirstName = DICTCONTACTINFO.Item("DT_FirstName")
	End If
	selectDateTime zipCode, dateToSelect
	provideContactInformation dictContactInfo
	submitContactInformation accType
	checkAppointmentConfirmation
	getRecordinCurrentModule dictContactInfo
	selectCreatedCustomerRecord dictContactInfo
	'Close appt for which call to office has been made
	If followUpScenario = "FollowUpAfterCallToOffice" Then
		addToContactInfo "DT_FirstName", tempFirstName
		addToContactInfo "DT_LastName",tempLastName
		addToContactInfo "DT_LastName",tempLastName
		addToCache "DT_CountLastFirstName",tempCountLastFirstName
		getRecordinCurrentModule dictContactInfo
		selectCreatedCustomerRecord dictContactInfo
		serviceComplete
	End If
End Function

Public Function checkDetailsWorkFlow()
	Details
	checkTabOptionsinDetailsPage
	Details
	verifyDataInDetailsTab dictContactInfo
	If CURRENT_VIEW = VIEW_PLATFORM Then
	   	Details
		'editDataInDemographicsTab dictupdatedContactInfo
		'Details
		verifyDatainDemographicsTab dictContactInfo
	End If
End Function



'########################
'Search Functions
'Only in Appointment Manager view
'Search by Associate Name - after assign associate
Public Function searchByAssociateName(inputData)
	Browser("SmartLobby").Page("SmartLobby").Image("AppointmentManager").Click
	Browser("SmartLobby").Page("SmartLobby").Frame("Frame_AppointmentManager").WebEdit("Associate").Set inputData
	Browser("SmartLobby").Page("SmartLobby").Frame("Frame_AppointmentManager").WebButton("Search").Click
	If Browser("SmartLobby").Page("SmartLobby").Frame("Frame_AppointmentManager").WbfGrid("ApptDataGrid").Exist(1) Then
    	rCount = Browser("SmartLobby").Page("SmartLobby").Frame("Frame_AppointmentManager").WbfGrid("ApptDataGrid").RowCount
    	colCount = Browser("SmartLobby").Page("SmartLobby").Frame("Frame_AppointmentManager").WbfGrid("ApptDataGrid").ColumnCount(1)
		For j=1 to colCount
			If UCase("Associate") = UCASE(Browser("SmartLobby").Page("SmartLobby").Frame("Frame_AppointmentManager").WbfGrid("ApptDataGrid").GetCellData(1,j)) Then
				colIndex = j
				Exit For
			End If
		Next
		searchFlag = True
		For k=2 to rCount
			currAssociate = Browser("SmartLobby").Page("SmartLobby").Frame("Frame_AppointmentManager").WbfGrid("ApptDataGrid").GetCellData(k,colIndex)
			If Instr(1, UCase(currAssociate), UCase(inputData)) = 0Then
				searchFlag = False
				Exit For
			End If
		Next
		If searchFlag= true Then
			logPass "Search with Associate Name is successful"
		Else
			logFail "Search with Associate Namefailed"
		End If
	Else
		logWarning "No Record found to validate the search criteria"
	End If
End Function

'Only in Appointment Manager view
'Search by Role
Public Function searchByRole(inputData)
	Browser("SmartLobby").Page("SmartLobby").Image("AppointmentManager").Click
	Browser("SmartLobby").Page("SmartLobby").Frame("Frame_AppointmentManager").WebList("Role").Select inputData
	Browser("SmartLobby").Page("SmartLobby").Frame("Frame_AppointmentManager").WebButton("Search").Click	
	If Browser("SmartLobby").Page("SmartLobby").Frame("Frame_AppointmentManager").WbfGrid("ApptDataGrid").Exist(1) Then
		rCount = Browser("SmartLobby").Page("SmartLobby").Frame("Frame_AppointmentManager").WbfGrid("ApptDataGrid").RowCount
		colCount = Browser("SmartLobby").Page("SmartLobby").Frame("Frame_AppointmentManager").WbfGrid("ApptDataGrid").ColumnCount(1)
		For j=1 to colCount
			If UCASE("Role") = UCASE(Browser("SmartLobby").Page("SmartLobby").Frame("Frame_AppointmentManager").WbfGrid("ApptDataGrid").GetCellData(1,j)) Then
				colIndex = j
				Exit For
			End If
		Next
		searchFlag = True
		For k=2 to rCount
			currRole = Browser("SmartLobby").Page("SmartLobby").Frame("Frame_AppointmentManager").WbfGrid("ApptDataGrid").GetCellData(k,colIndex)
			If UCase(currRole) <> UCase(inputData) Then
				searchFlag = False
				Exit For
			End If
		Next
		If searchFlag= true Then
			logPass "Search with role is successful"
		Else
			logFail "Search with role is failed"
		End If
	Else
		logWarning "No Record found to validate the search criteria"
	End If
End Function

'Only in Appointment Manager view
'Search by Customer
Public Function searchByCustomer(inputData)
	Browser("SmartLobby").Page("SmartLobby").Image("AppointmentManager").Click
	Browser("SmartLobby").Page("SmartLobby").Frame("Frame_AppointmentManager").WebEdit("Customer").Set inputData
	Browser("SmartLobby").Page("SmartLobby").Frame("Frame_AppointmentManager").WebButton("Search").Click
	If Browser("SmartLobby").Page("SmartLobby").Frame("Frame_AppointmentManager").WbfGrid("ApptDataGrid").Exist(1) Then
		rCount = Browser("SmartLobby").Page("SmartLobby").Frame("Frame_AppointmentManager").WbfGrid("ApptDataGrid").RowCount
    	colCount = Browser("SmartLobby").Page("SmartLobby").Frame("Frame_AppointmentManager").WbfGrid("ApptDataGrid").ColumnCount(1)
		For j=1 to colCount
			If UCASE("Customer") = UCASE(Browser("SmartLobby").Page("SmartLobby").Frame("Frame_AppointmentManager").WbfGrid("ApptDataGrid").GetCellData(1,j)) Then
				colIndex = j
				Exit For
			End If
		Next
		searchFlag = True
		For k=2 to rCount
			currCustomer = Browser("SmartLobby").Page("SmartLobby").Frame("Frame_AppointmentManager").WbfGrid("ApptDataGrid").GetCellData(k,colIndex)
			If Instr(1, UCase(currCustomer), UCase(inputData)) = 0Then
				searchFlag = False
				Exit For
			End If
		Next
		If searchFlag= true Then
			logPass "Search with role is successful"
		Else
			logFail "Search with role is failed"
		End If
	Else
		logWarning "No Record found to validate the search criteria"
	End If
End Function

'In Activity Monitoring tab
Public Function searchByName(inputData)
	Browser("SmartLobby").Page("SmartLobby").Image("ActivityMonitoring").Click
	Browser("SmartLobby").Page("SmartLobby").Frame("Frame_ActivityManager").WebEdit("Name").Set inputData
	Browser("SmartLobby").Page("SmartLobby").Frame("Frame_ActivityManager").WebButton("Search").Click
	rCount = Browser("SmartLobby").Page("SmartLobby").Frame("Frame_ActivityManager").WebTable("webTable_ActivityMonitor").RowCount
	If rCount >2 Then
		colCount = Browser("SmartLobby").Page("SmartLobby").Frame("Frame_ActivityManager").WebTable("webTable_ActivityMonitor").ColumnCount(1)
		For j=1 to colCount
			If UCASE("Customer") = UCASE(Browser("SmartLobby").Page("SmartLobby").Frame("Frame_ActivityManager").WebTable("webTable_ActivityMonitor").GetCellData(1,j)) Then
				colIndex = j
				Exit For
			End If
		Next
		searchFlag = True
		For k=3 to rCount			
				actualName = Browser("SmartLobby").Page("SmartLobby").Frame("Frame_ActivityManager").WebTable("webTable_ActivityMonitor").GetCellData(k,colIndex)
				If Instr(1,actualName,"ERROR: The specified cell does not exist.") = 0  Then
					If Instr(1, UCase(actualName), UCase(inputData)) = 0 Then
						searchFlag = False
						Exit For
					End If
				End If
		Next
		If searchFlag= true Then
			logPass "Search with Associate Name is successful"
		Else
			logFail "Search with Associate Namefailed"
		End If
	Else
		logWarning "No Record found to validate the search criteria"
	End If
End Function
'####################
