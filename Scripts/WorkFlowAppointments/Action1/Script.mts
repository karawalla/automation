Init()
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




