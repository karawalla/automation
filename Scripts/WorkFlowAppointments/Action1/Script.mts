﻿Init()

loginToScheduler tdGetUserName,tdGetPassword
goToLoginView tdGetViewName
goToModule tdGetModuleName
selectBranch  tdGetBranchType, tdGetBranchName,tdGetLocationType

selectAppointmentType tdGetAppointmentType
createAppointment tdGetAddDiscussionTopicDict, tdGetDateToSelect, tdGetAddContactInfoDict

validateAppointmentDetails tdGetDetailsFlag,tdGetAddContactInfoDict

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
validateAppointmentNoShow tdGetNoshowFlag, tdGetAddContactInfoDict

endTest










