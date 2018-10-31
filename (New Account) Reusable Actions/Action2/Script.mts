'-----------------------------------------------------------------------------------------------------------------------------------------------
'Action Name : Confirm_NewAccount
'Description : This action is to confirm new user account creation
'Creation Date : 29.10.2018
'Last modification date : None
'Assumptions /Effects : New user account is confirmed and created correctly
'Returns : Action returns 0 if everything is correct or returns ActionNumber + step number if error occurs
'-----------------------------------------------------------------------------------------------------------------------------------------------
Option Explicit
Reporter.Filter = rfDisableAll

' Click on Create Account button
If Browser("WebBrowser").Page("NewAccount").WebButton("CreateAccount").Exist(5) Then
	Browser("WebBrowser").Page("NewAccount").WebButton("CreateAccount").Click
Else
	fnReportStepEx "Fail", "Click on Create Account Button", "Create Account button is NOT displayed", Browser("WebBrowser"), "true"
	ExitActionIteration "Confirm_NewAccount.1"	
End If

If Browser("WebBrowser").Page("Dashboard").Exist(20) Then
	fnReportStepEx "Pass", "Click on Create Account Button", "Account is confirmed", Browser("WebBrowser"), "true"
Else
	fnReportStepEx "Fail", "Click on Create Account Button", "Account is NOT confirmed", Browser("WebBrowser"), "true"
	ExitActionIteration "Confirm_NewAccount.2"
End If

'Verify User Dashboard subpage
RunAction "Verify_UserDashboard [(Main) Reusable Actions]", oneIteration

ExitActionIteration "0"
