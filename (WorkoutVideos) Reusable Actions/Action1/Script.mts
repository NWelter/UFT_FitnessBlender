'Action Name : Verify_WorkoutVideos
'Description : This action is to verify that Workout Videos subpage is displayed and verify content
'Creation Date : 10.10.2018
'Last modification date : None
'Assumptions /Effects : Workout Videos subpage and content are correctly displayed
'Returns : Action returns 0 if everything is correct or returns ActionNumber + step number if error occurs
'-----------------------------------------------------------------------------------------------------------------------------------------------
Option Explicit
Reporter.Filter = rfDisableAll

Dim arrPageElements, arrCheckResults

' Verify Workout Videos subpage content
arrPageElements = Array(Browser("WebBrowser").Page("WorkoutVideos").WebElement("FreeWorkoutVideosHeader"),_ 
						Browser("WebBrowser").Page("WorkoutVideos").WebElement("FilterSearchVideoPanel"),_ 
						Browser("WebBrowser").Page("WorkoutVideos").WebElement("WorkoutVideosSection"))
						
arrCheckResults = fnCheckPageElements(arrPageElements)

If arrCheckResults(0) Then
	fnReportStepEx "Pass", "Click on Workout Videos subtab.", "Workout Videos subpage is displayed." & VbCrLf & "Current elements are displayed: " & arrCheckResults(1), Browser("WebBrowser"), "true"
Else 
	fnReportStepEx "Fail", "Click on Workout Videos subtab,", "Workout Videos subpage is NOT displayed." & VbCrLf & " Current elements are NOT displayed: " & arrCheckResults(2), Browser("WebBrowser"), "true"
	ExitActionIteration "Verify_WorkoutVideos.1"
End If

ExitActionIteration "0"
