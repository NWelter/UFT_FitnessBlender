'-----------------------------------------------------------------------------------------------------------------------------------------------
'Action Name : Check_UserDashboardHeader
'Description : This action is to compare text from header on User Dashboard subpage to <<strDashboardHeader>>
'Creation Date : 29.10.2018
'Last modification date : None
'Inputs: strDashboardHeader
'Assumptions /Effects : Text from header on User Dashboard subpage is compared to <<strDashboardHeader>>
'Returns : Action returns 0 if everything is correct or returns ActionNumber + step number if error occurs
'-----------------------------------------------------------------------------------------------------------------------------------------------
Option Explicit
Reporter.Filter = rfDisableAll

Dim strInvalidHeader

' Verify text from information header on User Dashboard subpage
If Browser("WebBrowser").Page("Dashboard").WebElement("ConfirmHeader").Exist(5) Then
	If fnContains(Browser("WebBrowser").Page("Dashboard").WebElement("ConfirmHeader"), "outertext", Parameter("strDashboardHeader")) Then
		fnReportStepEx "Pass", "Verify text from information header.", "Information header contains specified value: " &  Parameter("strDashboardHeader"), Browser("WebBrowser"), "true"
	Else
		strInvalidHeader = Browser("WebBrowser").Page("Dashboard").WebElement("ConfirmHeader").GetROProperty("outertext")
		fnReportStepEx "Fail", "Verify text from information header.", "Information header DOESN'T contain specified value: " &  Parameter("strDashboardHeader") & VbCrLf &_ 
		"Current text from information header is: " & strInvalidHeader, Browser("WebBrowser"), "true"
		ExitActionIteration "Check_UserDashboardHeader.1.1"
	End If
Else 
	fnReportStepEx "Fail", "Verify text from information header.", "Information header is NOT display", Browser("WebBrowser"), "true"
	ExitActionIteration "Check_UserDashboardHeader.1.2"
End If

ExitActionIteration "0"
