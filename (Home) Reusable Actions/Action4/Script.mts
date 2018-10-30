'------------------------------------------------------------------------------------------------------------
'Action Name: Hover_WorkoutsAndProgramsDropdown
'Description: This action is to verify that Workouts&Programs dropdown is expanded and current subtabs are displayed
'Creation Date: 08-10-2018
'Last modification date: <None>
'Assumptions / Effects: Workouts&Programs dropdown and current subtabs are displayed correctly
'Returns: Action return 0 if everything is correct or returns ActionNumber + step number if error occure
'------------------------------------------------------------------------------------------------------------

Option Explicit
Reporter.Filter = rfDisableAll

Dim arrPageElements, arrCheckResults

' Hover over Workouts&Programs dropdown
If Browser("WebBrowser").Page("AllPages").WebElement("FitnessBlenderHeader").WebElement("Workouts&ProgramsDropdown").Exist(5) Then
	Setting.WebPackage("ReplayType") = 2
	Browser("WebBrowser").Page("AllPages").WebElement("FitnessBlenderHeader").WebElement("Workouts&ProgramsDropdown").FireEvent("onmouseover")
	Setting.WebPackage("ReplayType") = 1
	
	' Verify current subtabs
	arrPageElements = Array(Browser("WebBrowser").Page("AllPages").WebElement("FitnessBlenderHeader").WebElement("Workouts&ProgramsDropdown").Link("WorkoutVideoSubtab"),_ 
				Browser("WebBrowser").Page("AllPages").WebElement("FitnessBlenderHeader").WebElement("Workouts&ProgramsDropdown").Link("WorkoutProgramsSubtab"),_ 
				Browser("WebBrowser").Page("AllPages").WebElement("FitnessBlenderHeader").WebElement("Workouts&ProgramsDropdown").Link("MealPlansSubtab"))

	arrCheckResults = fnCheckPageElements(arrPageElements)
	
	If arrCheckResults(0) Then
		fnReportStepEx "Pass", "Hover over Workouts&Programs dropdown. Verify current subtabs.", "Workouts&Programs dropdown is displayed." & VbCrLf &_ 
		"Current subtabs are displayed: " & arrCheckResults(1), Browser("WebBrowser"), "true"
	Else
		fnReportStepEx "Fail", "Hover over Workouts&Programs dropdown. Verify current subtabs.","Workouts&Programs dropdown is displayed." & VbCrLf &_
		"Current subtabs are NOT displayed." & arrCheckResults(2), Browser("WebBrowser"), "true"
		ExitActionIteration "Hover_WorkoutsAndProgramsDropdown.1.1"
	End If		
Else
	fnReportStepEx "Fail", "Hover over Workouts&Programs dropdown. Verify current subtabs.", "Workouts&Programs dropdown is NOT displayed", Browser("WebBrowser"), "true"
	ExitActionIteration "Hover_WorkoutsAndProgramsDropdown.1.2"
End If 

ExitActionIteration "0"

