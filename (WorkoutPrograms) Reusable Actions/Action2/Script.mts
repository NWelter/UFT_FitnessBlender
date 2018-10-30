'-----------------------------------------------------------------------------------------------------------------------------------------------
'Action Name : Navigate_WorkoutPrograms
'Description : This action is to verify that navigate to Workout Programs subpage is available
'Creation Date : 11.10.2018
'Last modification date : None
'Assumptions /Effects : Workout Programs subpage is displayed
'Returns : Action returns 0 if everything is correct or returns ActionNumber + step number if error occurs
'-----------------------------------------------------------------------------------------------------------------------------------------------
Option Explicit
Reporter.Filter = rfDisableAll

Dim strRunActionStatus

' Hover over Workouts&Programs dropdown
strRunActionStatus = "9999"
strRunActionStatus = RunAction ("Hover_WorkoutsAndProgramsDropdown [(Home) Reusable Actions]", oneIteration)
If (StrComp(strRunActionStatus, "0", 1) <> 0) Then
    fnReportStep "Fail", "Hover_WorkoutsAndProgramDropdown action failed" , "Returned value: " & strRunActionStatus , ""
    ExitActionIteration "Navigate_WorkoutPrograms.1"
End If

' Click on Workout Programs subtab
If Browser("WebBrowser").Page("AllPages").WebElement("FitnessBlenderHeader").WebElement("Workouts&ProgramsDropdown").Link("WorkoutProgramsSubtab").Exist  Then
	Browser("WebBrowser").Page("AllPages").WebElement("FitnessBlenderHeader").WebElement("Workouts&ProgramsDropdown").Link("WorkoutProgramsSubtab").Click
Else
	fnReportStepEx "Fail", "Click on Workout Programs subtab", "Workout Programs subtab is NOT displayed", Browser("WebBrowser"), "true"
	ExitActionIteration "Navigate_WorkoutPrograms.2"
End If

' Verify Workout Programs subpage content
RunAction "Verify_WorkoutPrograms", oneIteration

ExitActionIteration "0" 
