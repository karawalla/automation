Init()

<<<<<<< HEAD
loginToScheduler userName,password
goToLoginView viewName
goToModule moduleName
selectBranch  branchType, branchName

performAppointment tdGetApptType()


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
=======
loginToScheduler tdGetUserName,tdGetPassword
goToLoginView tdGetViewName
goToModule tdGetModuleName
selectBranch  tdGetBranchType, tdGetBranchName,tdGetLocationType

selectAppointmentType tdGetAppointmentType
createAppointment tdGetAddDiscussionTopicDict, tdGetDateToSelect, tdGetAddContactInfoDict

validateAppointmentDetails tdGetDetailsFlag,tdGetAddContactInfoDict
>>>>>>> 799fa35d0ecdc8a41351abe3f0fb25e3d8a8eb73

'Assign Associate
assignAssociateToAppointment tdGetAssignAssociateFlag, tdGetAssociateName

completeAppointmentService tdGetServiceCompleteFlag, tdGetAddContactInfoDict

'ToCancelAfterServiceComplete
cancelAppointmentAfterServiceComplete tdGetCancelAfterServiceCompleteFlag, tdGetAddContactInfoDict

'To HandOff
validateAppointmentHandOff tdGetHandoffFlag

'Follow up are possible only in Activity Monitoring tab for schedule workflow
followupAppointment tdGetFollowupFlag, tdGetReAppointmentType, tdGetAddDiscussionTopicDict, tdGetDateToSelect, tdGetAddContactInfoDict

'Cancel appointment
CancellationOfAppointment tdGetCancelAppointmentFlag, tdGetAddContactInfoDict

'No Show
<<<<<<< HEAD
If selectNoShow<> "" Then
	If selectNoShow = "TRUE" Then
		noShow(DICTCONTACTINFO)
	End If
End If
endTest
=======
validateAppointmentNoShow tdGetNoshowFlag, tdGetAddContactInfoDict

endTest










>>>>>>> 799fa35d0ecdc8a41351abe3f0fb25e3d8a8eb73
