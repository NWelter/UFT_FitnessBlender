'-----------------------------------------------------------------------------------------------------------------------------------------------
'Action Name : Verify_WorkoutPrograms
'Description : This action is to verify that Workout Programs subpage is displayed and verify content
'Author: Natalia Welter
'Creation Date : 10.10.2018
'Last modification date : None
'Assumptions /Effects : Workout Programs subpage and content are correctly displayed
'Returns : Action returns 0 if everything is correct or returns ActionNumber + step number if error occurs
'-----------------------------------------------------------------------------------------------------------------------------------------------
Option Explicit
Reporter.Filter = rfDisableAll

Dim arrPageElements, arrCheckResults

' Verify Workout Programs subpage content
arrPageElements = Array(Browser("WebBrowser").Page("WorkoutPrograms").WebElement("WorkoutProgramsHeader"),_ 
						Browser("WebBrowser").Page("WorkoutPrograms").WebElement("FilterSearchProgramPanel"),_ 
						Browser("WebBrowser").Page("WorkoutPrograms").WebElement("ProgramsVideosSection"))
						
arrCheckResults = fnCheckPageElements(arrPageElements)

If arrCheckResults(0) Then
	fnReportStepEx "Pass", "Click on Workout Programs subtab. Verify Workout Programs subpage content.",_ 
	"Workout Programs subpage is displayed." & VbCrLf & "Current elements are displayed: " & arrCheckResults(1), Browser("WebBrowser"), "true"
Else 
	fnReportStepEx "Fail", "Click on Workout Programs subtab. Verify Workout Programs subpage content.",_ 
	"Workout Videos subpage is NOT displayed." & VbCrLf & " Current elements are NOT displayed: " & arrCheckResults(2), Browser("WebBrowser"), "true"
	ExitActionIteration "Verify_WorkoutPrograms.1"
End If

ExitActionIteration "0"
