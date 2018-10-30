'------------------------------------------------------------------------------------------------------------
'Action Name: FillRegisterForm_Join
'Description: This action is to verify that user is able to register to application
'Creation Date: 08-10-2018
'Last modification date: <None>
'Assumptions / Effects: User is registered succesfully
'Inputs: strFirstName, strLastName, strUsername, strPassword, strConfirmPassword
'Outputs: strUsernameRandomize
'Returns: Action return 0 if everything is correct or returns ActionNumber + step number if error occure
'------------------------------------------------------------------------------------------------------------

Option Explicit
Reporter.Filter = rfDisableAll

Dim strRandomEmail, strRandomUsername

strRandomUsername = Parameter("strUsername") & fnRandomString()
Parameter("strUsernameRandomize") = strRandomUsername
strRandomEmail = fnRandomString & fnRandomNumber(3,7) & "@wp.pl"

' Set First Name
If Parameter("strFirstName") <> "" Then
	If Browser("WebBrowser").Page("Join").WebEdit("FirstName").Exist(5) Then
		If fnSet(Browser("WebBrowser").Page("Join").WebEdit("FirstName"), Parameter("strFirstName")) Then
		fnReportStepEx "Pass", "Set First Name as: '" & Parameter("strFirstName") & "'", "First Name field is set as '" & Parameter("strFirstName") & "'", Browser("WebBrowser"), "true"
		Else
			fnReportStepEx "Fail", "Set First Name as '" & Parameter("strFirstName") & "'", "First Name field is NOT set as '" & Parameter("strFirstName") & "'", Browser("WebBrowser"), "true"
			ExitActionIteration "FillRegisterForm_Join.1.1"			
		End If
	Else 
		fnReportStepEx "Fail", "Set First Name as '" & Parameter("strFirstName") & "'", "First Name field NOT exist", Browser("WebBrowser"), "true"
		ExitActionIteration "FillRegisterForm_Join.1.2"
	End If
End If

' Set Last Name
If Parameter("strLastName") <> "" Then
	If Browser("WebBrowser").Page("Join").WebEdit("LastName").Exist(5) Then
		If fnSet(Browser("WebBrowser").Page("Join").WebEdit("LastName"), Parameter("strLastName")) Then
		fnReportStepEx "Pass", "Set Last Name as: '" & Parameter("strFirstName") & "'", "Last Name field is set as '" & Parameter("strLastName") & "'", Browser("WebBrowser"), "true"
		Else
			fnReportStepEx "Fail", "Set Last Name as '" & Parameter("strLastName") & "'", "Last Name field is NOT set as '" & Parameter("strLastName") & "'", Browser("WebBrowser"), "true"
			ExitActionIteration "FillRegisterForm_Join.2.1"			
		End If
	Else 
		fnReportStepEx "Fail", "Set Last Name as '" & Parameter("strLastName") & "'", "Last Name field NOT exist", Browser("WebBrowser"), "true"
		ExitActionIteration "FillRegisterForm_Join.2.2"
	End If
End If

' Set E-mail
	If Browser("WebBrowser").Page("Join").WebEdit("Email").Exist(5) Then
		If fnSet(Browser("WebBrowser").Page("Join").WebEdit("Email"), strRandomEmail) Then
			fnReportStepEx "Pass", "Set E-mail as '" & strRandomEmail & "'", "E-mail field is set as '" & strRandomEmail & "'", Browser("WebBrowser"), "true"
		Else
			fnReportStepEx "Fail", "Set E-mail as '" & strRandomEmail & "'", "E-mail field is NOT set as '" & strRandomEmail & "'", Browser("WebBrowser"), "true"
			ExitActionIteration "FillRegisterForm_Join.3.1"
		End If
	Else
		fnReportStepEx "Fail", "Set E-mail as '" & strRandomEmail & "'", "E-mail field is NOT exist", Browser("WebBrowser"), "true"
		ExitActionIteration "FillRegisterForm_Join.3.2"
	End If


' Set Username
If Parameter("strUsername") <> "" Then
	If Browser("WebBrowser").Page("Join").WebEdit("Username").Exist(5) Then
		If fnSet(Browser("WebBrowser").Page("Join").WebEdit("Username"), strRandomUsername) Then
			fnReportStepEx "Pass", "Set Username as '" & strRandomUsername & "'", "Username field is set as '" & strRandomUsername & "'", Browser("WebBrowser"), "true"
		Else
			fnReportStepEx "Fail", "Set Username as '" & strRandomUsername & "'", "Username field is NOT set as '" & strRandomUsername & "'", Browser("WebBrowser"), "true"
			ExitActionIteration "FillRegisterForm_Join.4.1"
		End If
	Else
		fnReportStepEx "Fail", "Set Username as '" & strRandomUsername & "'", "Username field is NOT exist", Browser("WebBrowser"), "true"
		ExitActionIteration "FillRegisterForm_Join.4.2"
	End If
End If

' Set Password
Browser("WebBrowser").Page("Join").WebEdit("Password").SetSecure Parameter("strPassword")

' Check if Password field is set
If fnWaitTillValueIsNotEmpty(Browser("WebBrowser").Page("Join").WebEdit("Password")) Then		
	fnReportStepEx "Pass", "Set Password as '" & Parameter("strPassword") & "'", "Password field is set as '" & Parameter("strPassword") & "'", Browser("WebBrowser"), "true"
Else
	fnReportStepEx "Fail", "Set Password as '" & Parameter("strPassword") & "'", "Password field is empty", Browser("WebBrowser"), "true"
	ExitActionIteration "FillRegisterForm_Join.5"		
End If

' Set Confirm Password
Browser("WebBrowser").Page("Join").WebEdit("ConfirmPassword").SetSecure Parameter("strConfirmPassword")

' Check if Confirm Password field is set
If fnWaitTillValueIsNotEmpty(Browser("WebBrowser").Page("Join").WebEdit("ConfirmPassword")) Then		
	fnReportStepEx "Pass", "Set Confirm Password as '" & Parameter("strConfirmPassword") & "'", "Confirm Password field is set as '" & Parameter("strConfirmPassword") & "'", Browser("WebBrowser"), "true"
Else
	fnReportStepEx "Fail", "Set Confirm Password as '" & Parameter("strConfirmPassword") & "'", "Confirm Password field is empty", Browser("WebBrowser"), "true"
	ExitActionIteration "FillRegisterForm_Join.6"		
End If

' Select reCAPTCHA
If Browser("WebBrowser").Page("Join").WebElement("Recaptcha").Exist(5) Then
	Browser("WebBrowser").Page("Join").WebElement("Recaptcha").Click
	Wait 5
	fnReportStepEx "Pass", "Select reCAPTCHA", "reCAPTCHA is selected", Browser("WebBrowser"), "true"
Else
	fnReportStepEx "Fail", "Select reCAPTCHA", "reCAPTCHA is NOT displayed", Browser("WebBrowser"), "true"
		ExitActionIteration "FillRegisterForm_Join.7.2"
End If

' Click on Join button
If Browser("WebBrowser").Page("Join").WebButton("JoinButton").Exist(5) Then
	 Browser("WebBrowser").Page("Join").WebButton("JoinButton").Click
	 fnReportStepEx "Pass", "Click on Join button", "Join button is displayed", Browser("WebBrowser"), "true"
Else
	fnReportStepEx "Fail", "Click on Join button", "Join button is NOT displayed", Browser("WebBrowser"), "true"
	ExitActionIteration "FillRegisterForm_Join.8"
End If

ExitActionIteration "0"
