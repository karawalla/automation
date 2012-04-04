
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






