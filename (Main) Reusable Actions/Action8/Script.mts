'-----------------------------------------------------------------------------------------------------------------------------------------------
'Action Name : Navigate_SignIn
'Description : This action is to verify that navigate to Sign In subpage is available
'Creation Date : 11.10.2018
'Last modification date : None
'Assumptions /Effects : Sign In subpage is displayed
'Returns : Action returns 0 if everything is correct or returns ActionNumber + step number if error occurs
'-----------------------------------------------------------------------------------------------------------------------------------------------
Option Explicit
Reporter.Filter = rfDisableAll

Dim strRunActionStatus

' Hover over My Fitness dropdown
strRunActionStatus = "9999"
strRunActionStatus = RunAction ("Hover_MyFitness [(Home) Reusable Actions]", oneIteration)
If (StrComp(strRunActionStatus, "0", 1) <> 0) Then
    fnReportStep "Fail", "Hover_MyFitness action failed" , "Returned value: " & strRunActionStatus , ""
    ExitActionIteration "Navigate_SignIn.1"
End If

' Click on Sign In button
If Browser("WebBrowser").Page("AllPages").WebElement("FitnessBlenderHeader").WebElement("MyFitnessMenu").Link("SignIn").Exist(5) Then
	Browser("WebBrowser").Page("AllPages").WebElement("FitnessBlenderHeader").WebElement("MyFitnessMenu").Link("SignIn").Click
Else
	fnReportStepEx "Fail", "Click on Sign In button", "Sign In button is NOT displayed", Browser("WebBrowser"), "true"
	ExitActionIteration "Navigate_SignIn.2"	
End If

' Verify Sign In subpage content
RunAction "Verify_SignIn", oneIteration

ExitActionIteration "0"
