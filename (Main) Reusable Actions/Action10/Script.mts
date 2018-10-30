'-----------------------------------------------------------------------------------------------------------------------------------------------
'Action Name : Verify_UserDashboard
'Description : This action is to verify that User Dashboard subpage is displayed and verify content
'Creation Date : 29.10.2018
'Last modification date : None
'Assumptions /Effects : User Dashboard subpage and content are correctly displayed
'Returns : Action returns 0 if everything is correct or returns ActionNumber + step number if error occurs
'-----------------------------------------------------------------------------------------------------------------------------------------------
Option Explicit
Reporter.Filter = rfDisableAll

Dim arrPageElements, arrCheckResults

' Verify User Dashboard subpage content
arrPageElements = Array(Browser("WebBrowser").Page("Dashboard").WebElement("DashboardPanel"))

arrCheckResults = fnCheckPageElements(arrPageElements)

If arrCheckResults(0) Then
	fnReportStepEx "Pass", "Verify User Dashboard subpage content.", "User Dashboard subpage is displayed." & VbCrLf & "Current elements are displayed: " & arrCheckResults(1), Browser("WebBrowser"), "true"
Else 
	fnReportStepEx "Fail", "Verify User Dashboard subpage content.", "User Dashboard subpage is NOT displayed." & VbCrLf & "Current elements are displayed: " & arrCheckResults(2), Browser("WebBrowser"), "true"	
	ExitActionIteration "Verify_UserDashboard.1"
End If

ExitActionIteration "0"
