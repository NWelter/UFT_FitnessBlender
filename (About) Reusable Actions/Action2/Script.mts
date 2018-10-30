'-----------------------------------------------------------------------------------------------------------------------------------------------
'Action Name : Navigate_About
'Description : This action is to verify that navigate to About subpage is available
'Creation Date : 12.10.2018
'Last modification date : None
'Assumptions /Effects : About subpage is displayed
'Returns : Action returns 0 if everything is correct or returns ActionNumber + step number if error occurs
'-----------------------------------------------------------------------------------------------------------------------------------------------
Option Explicit
Reporter.Filter = rfDisableAll

' Click on About tab
If Browser("WebBrowser").Page("AllPages").WebElement("FitnessBlenderHeader").WebElement("About").Exist(5) Then
	Browser("WebBrowser").Page("AllPages").WebElement("FitnessBlenderHeader").WebElement("About").Click
Else
	fnReportStepEx "Fail", "Click on About tab", "About tab is NOT displayed", Browser("WebBrowser"), "true"
	ExitActionIteration "Navigate_About.1"
End If

' Verify About subpage
RunAction "Verify_About", oneIteration

ExitActionIteration "0"
