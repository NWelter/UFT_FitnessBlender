'-----------------------------------------------------------------------------------------------------------------------------------------------
'Action Name : Navigate_Cart
'Description : This action is to verify that navigate to Cart subpage is available
'Creation Date : 11.10.2018
'Last modification date : None
'Assumptions /Effects : Cart subpage is displayed
'Returns : Action returns 0 if everything is correct or returns ActionNumber + step number if error occurs
'-----------------------------------------------------------------------------------------------------------------------------------------------
Option Explicit
Reporter.Filter = rfDisableAll

' Click on Shopping Bag icon
If Browser("WebBrowser").Page("AllPages").WebElement("FitnessBlenderHeader").WebElement("ShoppingbagIcon").Exist(5) Then
	Browser("WebBrowser").Page("AllPages").WebElement("FitnessBlenderHeader").WebElement("ShoppingbagIcon").Click
Else
	fnReportStepEx "Fail", "Click on Shopping Bag icon on navbar header", "Shopping Bag icon is NOT displayed", Browser("WebBrowser"), "true"
	ExitActionIteration "Navigate_Cart.1"
End If

' Verify Cart subpage content
RunAction "Verify_Cart", oneIteration

ExitActionIteration "0"
