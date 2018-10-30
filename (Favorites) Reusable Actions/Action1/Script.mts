'-----------------------------------------------------------------------------------------------------------------------------------------------
'Action Name : Navigate_Favorites
'Description : This action is to verify that navigate to Favorites subpage is available
'Creation Date : 11.10.2018
'Last modification date : None
'Assumptions /Effects : Favorites subpage is displayed
'Returns : Action returns 0 if everything is correct or returns ActionNumber + step number if error occurs
'-----------------------------------------------------------------------------------------------------------------------------------------------
Option Explicit
Reporter.Filter = rfDisableAll

'Click on Favorites link
If Browser("WebBrowser").Page("AllPages").WebElement("FitnessBlenderHeader").WebElement("MyFitnessMenuUser").Link("Favorites").Exist(5) Then
	Browser("WebBrowser").Page("AllPages").WebElement("FitnessBlenderHeader").WebElement("MyFitnessMenuUser").Link("Favorites").Click
Else
	fnReportStepEx "Fail", "Click on Favorites link", "Favorites link is NOT displayed", Browser("WebBrowser"), "true"
	ExitActionIteration "Navigate_Favorites.1"
End If

'Verify Favorites subpage content
RunAction "Verify_Favorites", oneIteration

ExitActionIteration "0"




