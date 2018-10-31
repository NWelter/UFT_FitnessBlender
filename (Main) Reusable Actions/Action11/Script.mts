'------------------------------------------------------------------------------------------------------------
'Action Name: SignOut
'Description: This action is to verify that login user is able to log out from application and redirected to Home page
'Creation Date: 29-10-2018
'Last modification date: <None>
'Assumptions / Effects: User is log out succesfully
'Returns: Action return 0 if everything is correct or returns ActionNumber + step number if error occure
'------------------------------------------------------------------------------------------------------------

Option Explicit
Reporter.Filter = rfDisableAll

Dim strRunActionStatus

' Hover over My Fitness dropdown
strRunActionStatus = "9999"
strRunActionStatus = RunAction ("Hover_MyFitnessLoggedUser [(Home) Reusable Actions]", oneIteration)
If (StrComp(strRunActionStatus, "0", 1) <> 0) Then
    fnReportStep "Fail", "Hover_MyFitnessLoggedUser action failed" , "Returned value: " & strRunActionStatus , ""
    ExitActionIteration "SignOut.1"
End If

' Click on Sign Out button
If Browser("WebBrowser").Page("AllPages").WebElement("FitnessBlenderHeader").WebElement("MyFitnessMenuUser").Link("SignOut").Exist(5) Then
	Browser("WebBrowser").Page("AllPages").WebElement("FitnessBlenderHeader").WebElement("MyFitnessMenuUser").Link("SignOut").Click
	'fnReportStepEx "Pass", "Click on Sign Out button", "Sign Out button is displayed", Browser("WebBrowser"), "true"
Else
	fnReportStepEx "Fail", "Click on Sign Out button", "Sign Out button is NOT displayed", Browser("WebBrowser"), "true"
	ExitActionIteration "SignOut.2"
End If

' Verify redirection to Home page
If Browser("WebBrowser").Page("Home").Exist(20) Then
	fnReportStepEx "Pass", "Click on Sign Out button", "Home page is displayed.", Browser("WebBrowser"), "true"
Else
	fnReportStepEx "Fail", "Click on Sign Out button", "Home page is NOT displayed.", Browser("WebBrowser"), "true"
End If

ExitActionIteration "0"


