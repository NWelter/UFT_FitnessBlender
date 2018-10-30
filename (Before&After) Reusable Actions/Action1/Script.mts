'-----------------------------------------------------------------------------------------------------------------------------------------------
'Action Name : Verify_Before&After
'Description : This action is to verify that Before&After subpage is displayed and verify content
'Creation Date : 11.10.2018
'Last modification date : None
'Assumptions /Effects : Before&After subpage and content are correctly displayed
'Returns : Action returns 0 if everything is correct or returns ActionNumber + step number if error occurs
'-----------------------------------------------------------------------------------------------------------------------------------------------
Option Explicit
Reporter.Filter = rfDisableAll

Dim arrPageElements, arrCheckResults

' Verify Before&After subpage content
arrPageElements = Array(Browser("WebBrowser").Page("Before&After").WebElement("Before&AfterHeader"),_ 
						Browser("WebBrowser").Page("Before&After").WebElement("Before&AfterArticlesSection"))

arrCheckResults = fnCheckPageElements(arrPageElements)

If arrCheckResults(0) Then
	fnReportStepEx "Pass", "Click on Before&After subtab. Verify Before&After subpage content", "Before&After subpage is displayed." & VbCrLf & "Current elements are displayed: " & arrCheckResults(1), Browser("WebBrowser"), "true"
Else 
	fnReportStepEx "Fail", "Click on Before&After subtab. Verify Before&After subpage content", "Before&After subpage is NOT displayed." & VbCrLf & " Current elements are NOT displayed: " & arrCheckResults(2), Browser("WebBrowser"), "true"
	ExitActionIteration "Verify_Before&After.1"
End If

ExitActionIteration "0"
