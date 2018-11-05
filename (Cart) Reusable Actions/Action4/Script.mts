'-----------------------------------------------------------------------------------------------------------------------------------------------
'Action Name : Check_Item
'Description : This action is to check if selected items are presented in Cart
'Creation Date : 05.11.2018
'Last modification date : None
'Inputs: strItemLink
'Assumptions /Effects : Items are presented in Cart correctly
'Returns : Action returns 0 if everything is correct or returns ActionNumber + step number if error occurs
'-----------------------------------------------------------------------------------------------------------------------------------------------
Option Explicit
Reporter.Filter = rfDisableAll

' Check if selected item link <<strItemLink>> is presented in Cart
If Browser("WebBrowser").Page("Cart").WebElement("ShoppingCartPanel").Link("ItemLink").Exist(5) Then
	Browser("WebBrowser").Page("Cart").WebElement("ShoppingCartPanel").Link("ItemLink").SetTOProperty "href", Parameter ("strItemLink")
Else
	fnReportStepEx "Fail", "Check if selected item link: " & Parameter ("strItemLink") & " is presented in Cart",_ 
	"Item link is NOT displayed", Browser("WebBrowser"), "true"
	ExitActionIteration "Check_Item.1"
End If

If Browser("WebBrowser").Page("Cart").WebElement("ShoppingCartPanel").Link("ItemLink").Exist(5) Then
	Browser("WebBrowser").Page("Cart").WebElement("ShoppingCartPanel").Link("ItemLink").Object.scrollIntoView
	fnReportStepEx "Pass", "Check if selected item link: " & Parameter ("strItemLink") & " is presented in Cart",_ 
	"Selected item is presented in Cart", Browser("WebBrowser"), "true"
Else
	fnReportStepEx "Fail", "Check if selected item link: " & Parameter ("strItemLink") & " is presented in Cart",_ 
	"Selected item is NOT presented in Cart", Browser("WebBrowser"), "true"
	ExitActionIteration "Check_Item.2"
End If
	
ExitActionIteration "0"
	
