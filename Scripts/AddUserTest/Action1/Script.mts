Init()
'Login ot scheduler as master admin
loginToScheduler userName, password

'go to Lobby View
goToLobbyView()

'go to user manager page
goToUserManager()

assertUserNotExists "Demo User", "Verifying user does not exist"

addNewUser(addUserData)
assertTableContentExists Browser("SmartLobby").Page("SmartLobby").Frame("Frame_UserManager").WbfGrid("UserDataGrid"),retrieveFromCache("first name") & " " & retrieveFromCache("last name"), 1, "Verifying User exists"
endTest()




