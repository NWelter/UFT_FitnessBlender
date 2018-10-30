'Action Name : Verify_Fitness
'Description : This action is to verify that Fitness subpage is displayed and verify content
'Creation Date : 10.10.2018
'Last modification date : None
'Assumptions /Effects : Fitness subpage and content are correctly displayed
'Returns : Action returns 0 if everything is correct or returns ActionNumber + step number if error occurs
'-----------------------------------------------------------------------------------------------------------------------------------------------
Option Explicit
Reporter.Filter = rfDisableAll

Dim arrPageElements, arrCheckResults

' Verify Fitness subtab content
arrPageElements = Array(Browser("WebBrowser").Page("Fitness").WebElement("FitnessHeader"),_ 
						Browser("WebBrowser").Page("Fitness").WebElement("FitnessArticlesSection"))

arrCheckResults =fnCheckPageElements(arrPageElements)

If arrCheckResults(0) Then
	fnReportStepEx "Pass", "Click on Fitness subtab. Verify Fitness subpage content.", "Fitness subpage is displayed." & VbCrLf & "Current elements are displayed: " & arrCheckResults(1), Browser("WebBrowser"), "true"
Else 
	fnReportStepEx "Fail", "Click on Fitness subtab. Verify Fitness subpage content.", "Fitness subpage is NOT displayed." & VbCrLf & " Current elements are NOT displayed: " & arrCheckResults(2), Browser("WebBrowser"), "true"
	ExitActionIteration "Verify_Fitness.1"
End If

ExitActionIteration "0"
