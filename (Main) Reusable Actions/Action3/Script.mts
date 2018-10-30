'------------------------------------------------------------------------------------------------------------
'Action Name: SignIn
'Description: This action is to verify that user is able to sign in application
'Creation Date: 08-10-2018
'Last modification date: <None>
'Assumptions / Effects: User is logged succesfully
'Inputs: strUsername, strPassword
'Returns: Action return 0 if everything is correct or returns ActionNumber + step number if error occure
'------------------------------------------------------------------------------------------------------------

Option Explicit
Reporter.Filter = rfDisableAll

' Set Username
If Parameter("strUsername") <> "" Then
	If Browser("WebBrowser").Page("SignIn").WebElement("SignInBox").WebEdit("Username").Exist(5) Then	
		If fnSet(Browser("WebBrowser").Page("SignIn").WebElement("SignInBox").WebEdit("Username"), Parameter ("strUsername")) Then
			fnReportStepEx "Pass", "Set Username as '" & Parameter("strUsername") & "'", "Username field is set as '" & Parameter ("strUsername") & "'", Browser("WebBrowser"), "true"
		Else
			fnReportStepEx "Fail", "Set Username as '" & Parameter("strUsername") & "'", "Username field is NOT set as '" & Parameter ("strUsername") & "'", Browser("WebBrowser"), "true"
			ExitActionIteration "SignIn.1.1"			
		End If
	Else 
		fnReportStepEx "Fail", "Set Username as '" & Parameter("strUsername") & "'", "Username field NOT exist", Browser("WebBrowser"), "true"
		ExitActionIteration "SignIn.1.2"
	End  If
End If

' Set Password
Browser("WebBrowser").Page("SignIn").WebElement("SignInBox").WebEdit("Password").SetSecure Parameter("strPassword")

' Check if Password field is set
If 	fnWaitTillValueIsNotEmpty (Browser("WebBrowser").Page("SignIn").WebElement("SignInBox").WebEdit("Password")) Then
	fnReportStepEx "Pass", "Set Password as '" & Parameter("strPassword") & "'", "Password field is set as '" & Parameter ("strPassword") & "'", Browser("WebBrowser"), "true"
Else 
	fnReportStepEx "Fail", "Set Password as '" & Parameter("strPassword") & "'", "Password field is empty", Browser("WebBrowser"), "true"
	ExitActionIteration "SignIn.2"	
End If

' Click on Sign In button
If Browser("WebBrowser").Page("SignIn").WebElement("SignInBox").WebButton("SignInButton").Exist(5) Then
	Browser("WebBrowser").Page("SignIn").WebElement("SignInBox").WebButton("SignInButton").FireEvent("onclick")
Else 
	fnReportStepEx "Fail", "Click on Sign In button", "Sign In button is NOT displayed", Browser("WebBrowser"), "true"
	ExitActionIteration "SignIn.3"
End If

ExitActionIteration "0"

