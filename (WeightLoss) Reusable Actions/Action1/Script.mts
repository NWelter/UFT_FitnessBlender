'-----------------------------------------------------------------------------------------------------------------------------------------------
'Action Name : Verify_WeightLoss
'Description : This action is to verify that Weight Loss subpage is displayed and verify content
'Creation Date : 11.10.2018
'Last modification date : None
'Assumptions /Effects : Weight Loss subpage and content are correctly displayed
'Returns : Action returns 0 if everything is correct or returns ActionNumber + step number if error occurs
'-----------------------------------------------------------------------------------------------------------------------------------------------
Option Explicit
Reporter.Filter = rfDisableAll

Dim arrPageElements, arrCheckResults

' Verify Weight Loss subpage content
arrPageElements = Array(Browser("WebBrowser").Page("WeightLoss").WebElement("WeightLossHeader"),_ 
						Browser("WebBrowser").Page("WeightLoss").WebElement("WeightLossArticlesSection"))

arrCheckResults = fnCheckPageElements(arrPageElements)

If arrCheckResults(0) Then
	fnReportStepEx "Pass", "Click on Weight Loss subtab. Verify Weight Loss subpage content.", "Weight Loss subpage is displayed." & VbCrLf & "Current elements are displayed: " & arrCheckResults(1), Browser("WebBrowser"), "true"
Else 
	fnReportStepEx "Fail", "Click on Weight Loss subtab. Verify Weight Loss subpage content.", "Weight Loss subpage is NOT displayed." & VbCrLf & " Current elements are NOT displayed: " & arrCheckResults(2), Browser("WebBrowser"), "true"
	ExitActionIteration "Verify_WeightLoss.1"
End If

ExitActionIteration "0"
