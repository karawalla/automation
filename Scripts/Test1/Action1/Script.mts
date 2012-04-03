'########## Load the datatable, create results folder and initialize the test settings ##########
Init()

'############################# Scenario workflow ####################################'
'Login to scheduler as global admin
'go to Lobby view
'go to user manager view
'add new user
'verify new user has been added


loginToScheduler tdGetUserName(), tdGetPassword()

goToLoginView tdGetViewName()

goToUserManager()

addNewUser(tdGetNewUserDataDict())

assertUserExists retrieveFromCache("new user"), "Verifying new user has been added."

'############################# Clean up the test resources ####################################'
endTest()

