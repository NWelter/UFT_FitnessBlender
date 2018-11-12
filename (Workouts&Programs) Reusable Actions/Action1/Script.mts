'-----------------------------------------------------------------------------------------------------------------------------------------------
'Action Name : Verify_WorkoutsAndPrograms
'Description : This action is to verify that Workouts&Programs subpage is displayed and verify content
'Creation Date : 10.10.2018
'Last modification date : None
'Assumptions /Effects : Workouts&Programs subpage and content are correctly displayed
'Returns : Action returns 0 if everything is correct or returns ActionNumber + step number if error occurs
'-----------------------------------------------------------------------------------------------------------------------------------------------
Option Explicit
Reporter.Filter = rfDisableAll

Dim arrPageElementsTop, arrPageElementsMiddle, arrPageElementsBottom, arrCheckResults
Dim blnElementsAreDisplayed : blnElementsAreDisplayed = True

' Verify Workouts&Programs subpage top section content
arrPageElementsTop = Array (Browser("WebBrowser").Page("Workouts&Programs").WebElement("NewestProgramsHeader"),_ 
						Browser("WebBrowser").Page("Workouts&Programs").WebElement("NewestProgramsSection"),_ 
						Browser("WebBrowser").Page("Workouts&Programs").WebElement("NewestWorkoutVideoHeader"),_ 
						Browser("WebBrowser").Page("Workouts&Programs").WebElement("NewestWorkoutVideosSection"))
						
arrCheckResults = fnCheckPageElements(arrPageElementsTop)

If arrCheckResults(0) Then
	Browser("WebBrowser").Page("Workouts&Programs").WebElement("NewestWorkoutVideoHeader").Object.scrollIntoView
	fnReportStepEx "Pass", "Click on Workouts&Programs dropdown. Verify top section content.", "Workouts&Programs subpage is displayed." & VbCrLf & "Current sections are available: " & arrCheckResults(1), Browser("WebBrowser"), "true"
Else
	Browser("WebBrowser").Page("Workouts&Programs").WebElement("NewestWorkoutVideoHeader").Object.scrollIntoView
	fnReportStepEx "Fail", "Click on Workouts&Programs dropdown. Verify top section content.", "Workouts&Programs subpage is NOT displayed." & VbCrLf & "Current sections are NOT available: " & arrCheckResults(2), Browser("WebBrowser"), "true"
	blnElementsAreDisplayed = False
End If

' Verify Workouts&Programs subpage middle section content
arrPageElementsMiddle = Array (Browser("WebBrowser").Page("Workouts&Programs").WebElement("MealPlansHeader"),_ 
						Browser("WebBrowser").Page("Workouts&Programs").WebElement("MealPlansSection"),_
						Browser("WebBrowser").Page("Workouts&Programs").WebElement("BestFatLossProgramsHeader"),_ 
						Browser("WebBrowser").Page("Workouts&Programs").WebElement("BestFatLossProgramsSection"))
						
arrCheckResults = fnCheckPageElements(arrPageElementsMiddle)

If arrCheckResults(0) Then
	Browser("WebBrowser").Page("Workouts&Programs").WebElement("BestFatLossProgramsHeader").Object.scrollIntoView
	fnReportStepEx "Pass", "Click on Workouts&Programs dropdown. Verify middle section content.", "Workouts&Programs subpage is displayed." & VbCrLf & "Current sections are available: " & arrCheckResults(1), Browser("WebBrowser"), "true"
Else
	Browser("WebBrowser").Page("Workouts&Programs").WebElement("BestFatLossProgramsHeader").Object.scrollIntoView
	fnReportStepEx "Fail", "Click on Workouts&Programs dropdown. Verify middle section content.", "Workouts&Programs subpage is NOT displayed." & VbCrLf & "Current sections are NOT available: " & arrCheckResults(2), Browser("WebBrowser"), "true"
	blnElementsAreDisplayed = False	
End If

' Verify Workouts&Programs subpage bottom section content
arrPageElementsBottom = Array (Browser("WebBrowser").Page("Workouts&Programs").WebElement("HIITWorkoutsHeader"),_ 
						Browser("WebBrowser").Page("Workouts&Programs").WebElement("HIITWorkoutsSection"),_ 
						Browser("WebBrowser").Page("Workouts&Programs").WebElement("StrengthWorkoutsHeader"),_ 
						Browser("WebBrowser").Page("Workouts&Programs").WebElement("StrengthWorkoutsSection"),_ 
						Browser("WebBrowser").Page("Workouts&Programs").WebElement("BeginnerWorkoutsHeader"),_ 
						Browser("WebBrowser").Page("Workouts&Programs").WebElement("BeginnerWorkoutsSection"))
						
arrCheckResults = fnCheckPageElements(arrPageElementsBottom)

If arrCheckResults(0) Then
	Browser("WebBrowser").Page("Workouts&Programs").WebElement("StrengthWorkoutsHeader").Object.scrollIntoView
	fnReportStepEx "Pass", "Click on Workouts&Programs dropdown. Verify bottom section content.", "Workouts&Programs subpage is displayed." & VbCrLf & "Current sections are available: " & arrCheckResults(1), Browser("WebBrowser"), "true"
Else
	Browser("WebBrowser").Page("Workouts&Programs").WebElement("StrengthWorkoutsHeader").Object.scrollIntoView
	fnReportStepEx "Fail", "Click on Workouts&Programs dropdown. Verify bottom section content.", "Workouts&Programs subpage is NOT displayed." & VbCrLf & "Current sections are NOT available: " & arrCheckResults(2), Browser("WebBrowser"), "true"
	blnElementsAreDisplayed = False
End If

If NOT blnElementsAreDisplayed Then
	ExitActionIteration "Verify_WorkoutsAndPrograms.1"	
End If

ExitActionIteration "0"
						
	





