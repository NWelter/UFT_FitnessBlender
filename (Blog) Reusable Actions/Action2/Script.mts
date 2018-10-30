'-----------------------------------------------------------------------------------------------------------------------------------------------
'Action Name : Navigate_Blog
'Description : This action is to verify that navigate to Blog subpage is available
'Creation Date : 11.10.2018
'Last modification date : None
'Assumptions /Effects : Blog subpage is displayed
'Returns : Action returns 0 if everything is correct or returns ActionNumber + step number if error occurs
'-----------------------------------------------------------------------------------------------------------------------------------------------
Option Explicit
Reporter.Filter = rfDisableAll

' Click on Blog tab
If Browser("WebBrowser").Page("AllPages").WebElement("FitnessBlenderHeader").WebElement("Blog").Exist(5) Then
	Browser("WebBrowser").Page("AllPages").WebElement("FitnessBlenderHeader").WebElement("Blog").Click
Else
	fnReportStepEx "Fail", "Click on Blog tab", "Blog tab is NOT displayed", Browser("WebBrowser"), "true"
	ExitActionIteration "Navigate_Blog.1"
End If

' Verify Blog subpage content
RunAction "Verify_Blog", oneIteration

ExitActionIteration "0"
