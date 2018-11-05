'-----------------------------------------------------------------------------------------------------------------------------------------------
'Action Name : Verify_WorkoutsVideoDetails
'Description : This action is to verify Workout Video details subpage content
'Creation Date : 24.10.2018
'Author: Natalia Welter
'Last modification date : None
'Assumptions /Effects : Workout Video details subpage content is displayed correctly
'Returns : Action returns 0 if everything is correct or returns ActionNumber + step number if error occurs
'-----------------------------------------------------------------------------------------------------------------------------------------------
Option Explicit
Reporter.Filter = rfDisableAll

Dim arrPageElements, arrCheckResults

'Verify Workout Video details subpage content
arrPageElements = Array(Browser("WebBrowser").Page("WorkoutVideosDetails").WebElement("WorkoutDetailsPanel"),_ 
						Browser("WebBrowser").Page("WorkoutVideosDetails").WebButton("AddToFavorites"),_
						Browser("WebBrowser").Page("WorkoutVideosDetails").WebButton("AddToCalendar"),_ 
						Browser("WebBrowser").Page("WorkoutVideosDetails").WebElement("VideoHeader"),_ 
						Browser("WebBrowser").Page("WorkoutVideosDetails").WebElement("WorkoutVideoArticleBody"))

arrCheckResults = fnCheckPageElements(arrPageElements)

If arrCheckResults(0) Then
	fnReportStepEx "Pass", "Verify Workout Video details subpage content.",  "Current elements are displayed: " & arrCheckResults(1), Browser("WebBrowser"), "true"
Else 
	fnReportStepEx "Fail", "Verify Workout Video details subpage content.", "Current elements are NOT displayed: " & arrCheckResults(2), Browser("WebBrowser"), "true"	
	ExitActionIteration "Verify_WorkoutsVideoDetails.1"
End If

ExitActionIteration "0"
