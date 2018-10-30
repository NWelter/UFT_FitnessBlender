'-----------------------------------------------------------------------------------------------------------------------------------------------
'Action Name : Navigate_MealPlans
'Description : This action is to verify that navigate to Meal Plans subpage is available
'Creation Date : 11.10.2018
'Last modification date : None
'Assumptions /Effects : Meal Plans subpage is displayed
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
    ExitActionIteration "Navigate_MealPlans.1"
End If

' Click on Meal Plans subtab
If Browser("WebBrowser").Page("AllPages").WebElement("FitnessBlenderHeader").WebElement("Workouts&ProgramsDropdown").Link("MealPlansSubtab").Exist(5)  Then
	Browser("WebBrowser").Page("AllPages").WebElement("FitnessBlenderHeader").WebElement("Workouts&ProgramsDropdown").Link("MealPlansSubtab").Click
Else
	fnReportStepEx "Fail", "Click on Meal Plans subtab.", "Meal Plans subtab is NOT displayed", Browser("WebBrowser"), "true"
	ExitActionIteration "Navigate_MealPlans.2"
End If

'Verify Meal Plans subpage content
RunAction "Verify_MealPlans", oneIteration

ExitActionIteration "0"

