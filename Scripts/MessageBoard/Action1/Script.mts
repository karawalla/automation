Init()
'schedulerLogin tdGetUserName,tdGetPassword
'goToLoginView tdGetView
'goToCorporateNode

DataTable.SetCurrentRow(Environment("ActionIteration"))

schedulerLogin tdGetUserName,tdGetPassword
goToLoginView tdGetLoginView
goToModule tdGetModule
goToCorporateNode

If tdGetAddNewMessage<>"" Then
	If UCase(tdGetAddNewMessage) = "TRUE" Then
		addNewMessage tdGetTextMessage
		isMessageFound tdGetTextMessage
	End If
End If

If tdGetEditExistingMessage<>"" Then
	If UCase(tdGetEditExistingMessage) = "TRUE" Then
		editExistingMessage tdGetTextMessage,tdGetUpdatedTextMessage
		isMessageFound tdGetUpdatedTextMessage
	End If
End If

If tdGetDeleteExistingMessage<>"" Then
	If UCase(tdGetDeleteExistingMessage) = "TRUE" Then
		deleteExistingMessage tdGetTextMessage
		isMessageNotFound tdGetTextMessage
	End If
End If