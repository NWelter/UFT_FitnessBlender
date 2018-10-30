'-----------------------------------------------------------------------------------------------------------------------------------------------
'Action Name : Verify_Health
'Description : This action is to verify that Health subpage is displayed and verify content
'Creation Date : 11.10.2018
'Last modification date : None
'Assumptions /Effects : Health subpage and content are correctly displayed
'Returns : Action returns 0 if everything is correct or returns ActionNumber + step number if error occurs
'-----------------------------------------------------------------------------------------------------------------------------------------------
Option Explicit
Reporter.Filter = rfDisableAll

Dim arrPageElements, arrCheckResults

' Verify Health subpage content
arrPageElements = Array(Browser("WebBrowser").Page("Health").WebElement("HealthHeader"),_ 
						Browser("WebBrowser").Page("Health").WebElement("HealthArticlesSection"))

arrCheckResults = fnCheckPageElements(arrPageElements)

If arrCheckResults(0) Then
	fnReportStepEx "Pass", "Click on Health subtab. Verify Health subpage content.", "Health subpage is displayed." & VbCrLf & "Current elements are displayed: " & arrCheckResults(1), Browser("WebBrowser"), "true"
Else 
	fnReportStepEx "Fail", "Click on Health subtab. Verify Health subpage content.", "Health subpage is NOT displayed." & VbCrLf & " Current elements are NOT displayed: " & arrCheckResults(2), Browser("WebBrowser"), "true"
	ExitActionIteration "Verify_Health.1"
End If

ExitActionIteration "0"
