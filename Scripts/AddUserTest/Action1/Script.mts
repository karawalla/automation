Init()
'Login ot scheduler as master admin
loginToScheduler userName, password

'go to Lobby View
goToLobbyView()

'go to user manager page
goToUserManager()

'Verify 
assertUserNotExists "Demo User", "Verifying user does not exist"

addNewUser(addUserData)

assertUserExists retrieveFromCache("new user name")
endTest()




