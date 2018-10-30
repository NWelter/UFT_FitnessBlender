'-----------------------------------------------------------------------------------------------------------------------------------------------
'Action Name : Navigate_Community
'Description : This action is to verify that navigate to Community subpage is available
'Creation Date : 11.10.2018
'Last modification date : None
'Assumptions /Effects : Community subpage is displayed
'Returns : Action returns 0 if everything is correct or returns ActionNumber + step number if error occurs
'-----------------------------------------------------------------------------------------------------------------------------------------------
Option Explicit
Reporter.Filter = rfDisableAll

' Click on Community tab
If Browser("WebBrowser").Page("AllPages").WebElement("FitnessBlenderHeader").WebElement("Community").Exist(5) Then
	Browser("WebBrowser").Page("AllPages").WebElement("FitnessBlenderHeader").WebElement("Community").Click
Else
	fnReportStepEx "Fail", "Click on Community tab", "Community tab is NOT displayed", Browser("WebBrowser"), "true"
	ExitActionIteration "Navigate_Community.1"
End If

' Verify Community subpage
RunAction "Verify_Community", oneIteration

ExitActionIteration "0"

