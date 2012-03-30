x = CInt(weekDay(date))
Msgbox x
Msgbox weekDayName(x)

text = "test Msg board"
updatedText = "My sample Msg board to test updated workflowautomation1"
addNewMessage
verifyMessage(text)

editExistingMessage()
isMessageFound(updatedText)

deleteExistingMessage
isMessageNotFound(text)

Public Function deleteExistingMessage()
	Browser("SmartLobby").Page("SmartLobby").Link("MessageBoard").Click	
	rCount = Browser("SmartLobby").Page("SmartLobby").WbfGrid("MessageBoardGrid").RowCount	
	isFound = False
	For i=2 to rCount
		If UCase(Browser("SmartLobby").Page("SmartLobby").WbfGrid("MessageBoardGrid").GetCellData(i,1)) = UCase(text) Then			
			isFound = True
			Browser("SmartLobby").Page("SmartLobby").WbfGrid("MessageBoardGrid").Object.Rows(i).Click
			Set editLink = Browser("SmartLobby").Page("SmartLobby").WbfGrid("MessageBoardGrid").ChildItem(i,5,"Link",0)
			editLink.Click
			If Dialog("text:=Message from webpage").Exist(1) Then
				Dialog("text:=Message from webpage").WinButton("text:=OK").Click			
			End If
			Exit For
		End If
	Next
	If isFound Then
		logPass "Msg found in message board for deletion"
	Else
		logFatal "Message not Found in message board for deletion"
	End If
End Function


Public Function editExistingMessage()
	Browser("SmartLobby").Page("SmartLobby").Link("MessageBoard").Click	
	rCount = Browser("SmartLobby").Page("SmartLobby").WbfGrid("MessageBoardGrid").RowCount	
	isFound = False
	For i=2 to rCount
		If UCase(Browser("SmartLobby").Page("SmartLobby").WbfGrid("MessageBoardGrid").GetCellData(i,1)) = UCase(text) Then			
			isFound = True
			Browser("SmartLobby").Page("SmartLobby").WbfGrid("MessageBoardGrid").Object.Rows(i).Click
			Set editLink = Browser("SmartLobby").Page("SmartLobby").WbfGrid("MessageBoardGrid").ChildItem(i,4,"Link",0)
			editLink.Click
			'Custom Sync
			Browser("MessageBoardEditor").Page("MessageBoardEditor").Frame("Frame_DesignEditor").WebElement("TextEditor").Click
			Set shlObj = CreateObject("WScript.Shell")
			shlObj.SendKeys("^a")
			shlObj.SendKeys("{DEL}")
			shlObj.SendKeys(updatedText)
			Browser("MessageBoardEditor").Page("MessageBoardEditor").WebButton("MessagePreview").Click
			msgText = Browser("MessageBoardEditor").Page("MessageBoardEditor").WebElement("testMsgboard").GetROProperty("outertext")
			If UCase(msgText) = Ucase(updatedText) Then
				logPass "Message text has been populated"
			Else
				logFail "Message not populated"
			End If
			Browser("MessageBoardEditor").Page("MessageBoardEditor").WebEdit("DisplayStartDate").Set Date()
			Browser("MessageBoardEditor").Page("MessageBoardEditor").WebEdit("DisplayEndDate").Set Date()+10
			Browser("MessageBoardEditor").Page("MessageBoardEditor").WebButton("Save").Click
			Exit For
		End If
	Next
	If isFound Then
		logPass "Msg found in message board"
	Else
		logFail "Message not Found in message board"
	End If
End Function

Public Function addNewMessage()
	Browser("SmartLobby").Page("SmartLobby").Link("MessageBoard").Click
	Browser("SmartLobby").Page("SmartLobby").Frame("Frame_EM_MessageBoard").WebButton("Add Message").Click
	'Custom Sync
	Browser("MessageBoardEditor").Page("MessageBoardEditor").Frame("Frame_DesignEditor").WebElement("TextEditor").Click
	Set shlObj = CreateObject("WScript.Shell")
	shlObj.SendKeys(text)
	Browser("MessageBoardEditor").Page("MessageBoardEditor").WebButton("MessagePreview").Click
	msgText = Browser("MessageBoardEditor").Page("MessageBoardEditor").WebElement("testMsgboard").GetROProperty("outertext")
	If UCase(msgText) = Ucase(text) Then
		logPass "Message text has been populated"
	Else
		logFail "Message not populated"
	End If
	Browser("MessageBoardEditor").Page("MessageBoardEditor").WebEdit("DisplayStartDate").Set Date()
	Browser("MessageBoardEditor").Page("MessageBoardEditor").WebEdit("DisplayEndDate").Set Date()+10
	Browser("MessageBoardEditor").Page("MessageBoardEditor").WebButton("Save").Click
End Function

Public Function isMessageFound(inText)
   	Browser("SmartLobby").Page("SmartLobby").Image("Messages").Click
	rCount = Browser("SmartLobby").Page("SmartLobby").WbfGrid("MessageBoardGrid").RowCount	
	isFound = False
	For i=2 to rCount
		If UCase(Browser("SmartLobby").Page("SmartLobby").WbfGrid("MessageBoardGrid").GetCellData(i,1)) = UCase(inText) Then			
			isFound = True
			Exit For
		End If
	Next
	If isFound Then
		logPass "Msg found in message board"
	Else
		logFail "Message not Found in message board"
	End If
End Function

Public Function isMessageNotFound(inText)
   	Browser("SmartLobby").Page("SmartLobby").Image("Messages").Click
	rCount = Browser("SmartLobby").Page("SmartLobby").WbfGrid("MessageBoardGrid").RowCount	
	isFound = False
	For i=2 to rCount
		If UCase(Browser("SmartLobby").Page("SmartLobby").WbfGrid("MessageBoardGrid").GetCellData(i,1)) = UCase(inText) Then			
			isFound = True
			Exit For
		End If
	Next
	If isFound Then
		logFail "Msg found in message board"
	Else
		logPass "Message not Found in message board.Message deleted successfully"
	End If
End Function

Public Function checkCalendarSegment()
   If Browser("SmartLobby").Page("SmartLobby").Link("CalendarsnHolidays").Exist(1) Then
		Browser("SmartLobby").Page("SmartLobby").Link("CalendarsnHolidays").Click
		If Browser("SmartLobby").Page("SmartLobby").Frame("Frame_EM_CalendernHolidays").WebTable("CalendarName").Exist(1) Then
			If Browser("SmartLobby").Page("SmartLobby").Frame("Frame_EM_CalendernHolidays").WebElement("col_CalendarName").Exist(1) And _
				Browser("SmartLobby").Page("SmartLobby").Frame("Frame_EM_CalendernHolidays").WebElement("col_Description").Exist(1) And _
				Browser("SmartLobby").Page("SmartLobby").Frame("Frame_EM_CalendernHolidays").WebElement("col_shortName").Exist(1) Then
					logPass "Calender table is displayed with column names:CalendarName, shortName and Description"
			Else
					logFail "Calender table is not getting displayed correctly with expected column names"
			End If
		Else
			logFail "Calender table is not getting displayed"
		End If
	Else
   End If
End Function

Public Function checkTabOptionsinEMCorporateNode()
   If Browser("SmartLobby").Page("SmartLobby").Link("Services").Exist(1) Then
	   logPass "Currently in Corporate Node.Checking tab options availability"
	   If Browser("SmartLobby").Page("SmartLobby").Link("CalendarMapping").Exist(1) And _
		   Browser("SmartLobby").Page("SmartLobby").Link("CalendarsnHolidays").Exist(1) And _
		   Browser("SmartLobby").Page("SmartLobby").Link("CustomerTypeManager").Exist(1) And _
		   Browser("SmartLobby").Page("SmartLobby").Link("MessageBoard").Exist(1) Then
	 			logPass "Tab options displaying correctly"
		Else
			logFail "Tab options are not showing correctly"
	   End If
	Else
		logFail "Currently not in corporate node.Please check the application work flow"
   End If
End Function

Public Function goToCorporateNode()
   If Browser("SmartLobby").Page("SmartLobby").Frame("BankLocation_TreeView").Link("Bank of America Corporate").Exist(1)  Then
	   Browser("SmartLobby").Page("SmartLobby").Frame("BankLocation_TreeView").Link("Bank of America Corporate").Click
	   logPass "Naviagted to Corporate Node"
	Else
		logFatal "Corporate Node link not Found"
   End If
End Function






