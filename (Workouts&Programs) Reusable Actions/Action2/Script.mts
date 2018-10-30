'-----------------------------------------------------------------------------------------------------------------------------------------------
'Action Name : Navigate_WorkoutsAndPrograms
'Description : This action is to verify that navigate to Workouts&Programs subpage is available
'Creation Date : 11.10.2018
'Last modification date : None
'Assumptions /Effects : Workouts&Programs subpage is displayed
'Returns : Action returns 0 if everything is correct or returns ActionNumber + step number if error occurs
'-----------------------------------------------------------------------------------------------------------------------------------------------
Option Explicit
Reporter.Filter = rfDisableAll

' Click on Workouts&Programs dropdown
If Browser("WebBrowser").Page("AllPages").WebElement("FitnessBlenderHeader").Link("Workouts&ProgramsLink").Exist(5) Then
	Browser("WebBrowser").Page("AllPages").WebElement("FitnessBlenderHeader").Link("Workouts&ProgramsLink").Click
Else
	fnReportStepEx "Fail", "Click on Workouts&Programs dropdown", "Workouts&Programs dropdown is NOT displayed", Browser("WebBrowser"), "true"
	ExitActionIteration "Navigate_WorkoutsAndPrograms.1"
End If

'Verify Workouts&Programs subpage content
RunAction "Verify_WorkoutsAndPrograms", oneIteration

ExitActionIteration "0"

