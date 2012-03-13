'Test Data
DataTable.Import "WorkFlowAppointments.xls"
DataTable.SetCurrentRow Environment.Value("TestIteration")
Init()
'launchScheduler
'url = "http://qaweb2.ncr.com/SSMPortalBOA_Trunk/Login.aspx"

'schedulerLogin
loginData = Split(DataTable.Value("LoginData"),":")
userName = loginData(0)
password = loginData(1)
'goToLoginView
viewName = DataTable.Value("View")

'selectBranch
branchName = DataTable.Value("BranchName")
branchType = DataTable.Value("BranchType")

'goToModule
moduleName = DataTable.Value("Module")

'set workFlow
workFlow = DataTable.Value("WorkFlow", Global)
'Data for selectTopicforDiscussion
DiscussionTopic1 = DataTable.Value("DiscussionTopic1")
DiscussionTopic2 = DataTable.Value("DiscussionTopic2")
comments = "TestComments"
language = "English"
accType = DataTable.Value("AccType")

addToDiscussionTopic "DT_Topic1",DiscussionTopic1
addToDiscussionTopic "DT_Topic2",DiscussionTopic2
addToDiscussionTopic "DT_Comments", comments
addToDiscussionTopic "DT_Language", language
addToDiscussionTopic "DT_AccType", accType

'selectDateTime
zipCode = "75038"
dateToSelect = DataTable.Value("SelectDate")

'ProvideContactInformation
firstName = DataTable.Value("FirstName")
lastName = DataTable.Value("LastName")
eMail = "mouliayyala@gmail.com"
reEnterEmail ="mouliayyala@gmail.com"
contactNumber = "1234567890"
existingCustomer = "Yes"
businessName = "Test"
partyID ="123"
platForm = "COIN"
reminder_Both = "ON"
reminder_Phone = Empty
reminder_textMsg = Empty

addToContactInfo "DT_FirstName",firstName
addToContactInfo "DT_LastName",lastName
addToContactInfo "DT_Email",eMail
addToContactInfo "DT_ReEnterEmail",eMail
addToContactInfo "DT_ContactNumber",contactNumber
addToContactInfo "DT_ExistingCustomer",existingCustomer
addToContactInfo "DT_BusinessName",businessName
addToContactInfo "DT_PartyID",partyID
addToContactInfo "DT_PlatForm",platForm
addToContactInfo "DT_Reminder_Both",reminder_Both
addToContactInfo "DT_Reminder_Phone",reminder_Phone
addToContactInfo "DT_Reminder_TxtMsg",reminder_textMsg
addToContactInfo "DT_AccType",accType


setCurrentView viewName
setCurrentModule moduleName
setCurrentWorkFlow workFlow


setFollowupType DataTable.Value("AppointmentType")
serviceComp = DataTable.Value("ServiceComplete")
followUpScenario =DataTable.Value("FollowUpScenario")
toHandOff = DataTable.Value("ToHandOff", Global)
toCheckDetails = DataTable.Value("CheckDetails")
cancelAppt = DataTable.Value("CancelAppointment")
selectNoShow = DataTable.Value("SelectNoShow")
toAssignAssociate	= DataTable.Value("AssignAssociate")
associateName = DataTable.Value("AssociateName")
toCancelAfterServiceComplete = DataTable.Value("CheckCancelafterserviceComplete")
locationtype = DataTable.Value("LocationType")
addToCache "DT_Location",locationtype
'*************************************************************
schedulerLogin userName,password
goToLoginView viewName
goToModule moduleName
selectBranch  branchType, branchName

Select Case workFlow 'CURRENT_WORKFLOW
	Case WORKFLOW_SCHEDULE
		scheduleAppointmentWorkFlow
	Case WORKFLOW_WALKIN
		walkInWorkFlow
	Case WORKFLOW_RESCHEDULE
		reScheduleAppointmentWorkFlow
End Select

'To check Details
If toCheckDetails<>"" Then
	If toCheckDetails = "TRUE" Then
		checkDetailsWorkFlow
	End If
End If

'Assign Associate
If  toAssignAssociate<> "" Then
	If toAssignAssociate = "TRUE" Then
		goToAssignAssociate
		assignAssociate(associateName)
		verifyAssociateName()
	End If
End If

If serviceComp<>"" Then
	If serviceComp = "TRUE" Then
		Select Case workFlow 'CURRENT_WORKFLOW
			Case WORKFLOW_WALKIN
				'Do nothing
			Case WORKFLOW_RESCHEDULE
				apptCheckIn
			Case WORKFLOW_SCHEDULE
				apptCheckIn
		End Select
				serviceCompleteWorkFlow
	End If
End If

'ToCancelAfterServiceComplete
If toCancelAfterServiceComplete<>"" Then
	If toCancelAfterServiceComplete = "TRUE" Then
		CancelAfterServiceCompleteWorkflow
	End If
End If

'To HandOff
If toHandOff<>"" Then
	If toHandOff = "TRUE" Then
		If workFlow = WORKFLOW_WALKIN Then 'CURRENT_MODULE
			callToOffice
			handOff
		Else
			logFatal "Cannot handoff the selected appointment.Please check the workflow whether it is a walkin appointmentr or not"
		End If
	End If
End If

'Follow up are possible only in Activity Monitoring tab for schedule workflow
If followUpScenario<>"" Then
	If  moduleName = MODULE_ACTIVITYMONITORING AND workFlow <> WORKFLOW_WALKIN Then 'CURRENT_MODULE
		Select Case followUpScenario 'To add private property
			Case "FollowUpBeforeCheckIn"				
			Case "FollowUpAfterCheckIn"
				apptCheckIn
			Case "FollowUpAfterCallToOffice"
				apptCheckIn
				callToOffice
		End Select
		followUpWorkFlow
	Else
		logFatal "Followup appointments cannot be made from" & moduleName & "or it is a walkin appointment for which followup cannot be created " 'CURRENT_MODULE
	End If
End If

'Cancel appointment
If cancelAppt<> "" Then
	If cancelAppt = "TRUE" Then
		cancelAppointment
		verifyAppointmentCancellation()
	End If
End If

'No Show
If selectNoShow<> "" Then
	If selectNoShow = "TRUE" Then
		noShow(DICTCONTACTINFO)
	End If
End If
endTest
'##############################################################
'rValue = retrieveFromCache("DT_RecordNumber")
'If Browser("Smart Lobby").Page("Smart Lobby").Frame("Frame_AppointmentManager").Link("Details").Object.isDisabled Then
'	Browser("Smart Lobby").Page("Smart Lobby").Frame("Frame_AppointmentManager").WbfGrid("ApptDataGrid").Object.Rows(rValue-1).Click
'End If
'rValue = retrieveFromCache("DT_RecordNumber")
'If Browser("Smart Lobby").Page("Smart Lobby").Frame("Frame_ActivityManager").WebElement("Details").Object.isDisabled Then
'	Browser("Smart Lobby").Page("Smart Lobby").Frame("Frame_ActivityManager").WebTable("webTable_ActivityMonitor").Object.Rows(rValue-1).Click
'End If






