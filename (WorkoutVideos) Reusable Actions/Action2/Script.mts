'-----------------------------------------------------------------------------------------------------------------------------------------------
'Action Name : Navigate_WorkoutVideos
'Description : This action is to verify that navigate to Workout Videos subpage is available
'Creation Date : 11.10.2018
'Last modification date : None
'Assumptions /Effects : Workout Videos subpage is displayed
'Returns : Action returns 0 if everything is correct or returns ActionNumber + step number if error occurs
'-----------------------------------------------------------------------------------------------------------------------------------------------
Option Explicit
Reporter.Filter = rfDisableAll

Dim strRunActionStatus

'Hover over Workouts&Programs dropdown
strRunActionStatus = "9999"
strRunActionStatus = RunAction ("Hover_WorkoutsAndProgramsDropdown [(Home) Reusable Actions]", oneIteration)
If (StrComp(strRunActionStatus, "0", 1) <> 0) Then
    fnReportStep "Fail", "Hover_WorkoutsAndProgramDropdown action failed" , "Returned value: " & strRunActionStatus , ""
    ExitActionIteration "Navigate_WorkoutVideos.1"
End If

' Click on Workout Videos subtab
If Browser("WebBrowser").Page("AllPages").WebElement("FitnessBlenderHeader").WebElement("Workouts&ProgramsDropdown").Link("WorkoutVideoSubtab").Exist(5) Then
	Browser("WebBrowser").Page("AllPages").WebElement("FitnessBlenderHeader").WebElement("Workouts&ProgramsDropdown").Link("WorkoutVideoSubtab").Click
Else
	fnReportStepEx "Fail", "Click on Workout Videos subtab", "Workout Videos subtab is NOT displayed", Browser("WebBrowser"), "true"
	ExitActionIteration "Navigate_WorkoutVideos.2"
End If

'Verify Workout Videos subpage content
RunAction "Verify_WorkoutVideos", oneIteration

ExitActionIteration "0"

