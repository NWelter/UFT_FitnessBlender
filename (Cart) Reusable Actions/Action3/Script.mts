'-----------------------------------------------------------------------------------------------------------------------------------------------
'Action Name : Verify_Cart
'Description : This action is to verify that Cart subpage is displayed and verify content
'Creation Date : 02.11.2018
'Last modification date : None
'Assumptions /Effects : Cart subpage and content are correctly displayed
'Returns : Action returns 0 if everything is correct or returns ActionNumber + step number if error occurs
'-----------------------------------------------------------------------------------------------------------------------------------------------
Option Explicit
Reporter.Filter = rfDisableAll

Dim arrPageElements, arrCheckResults

' Verify Cart subpage content
arrPageElements = Array(Browser("WebBrowser").Page("Cart").WebElement("ShoppingCartHeader"),_ 
						Browser("WebBrowser").Page("Cart").WebElement("ShoppingCartPanel"),_ 
						Browser("WebBrowser").Page("Cart").WebElement("ShoppingCartPanel").WebElement("CartPanelHeader"),_ 
						Browser("WebBrowser").Page("Cart").WebElement("ShoppingCartPanel").Link("ItemHeader"),_ 
						Browser("WebBrowser").Page("Cart").WebElement("ShoppingCartPanel").WebElement("RemoveIcon"),_ 
						Browser("WebBrowser").Page("Cart").WebElement("ShoppingCartPanel").WebElement("Price"),_ 
						Browser("WebBrowser").Page("Cart").WebElement("ShoppingCartPanel").WebElement("OrderTotal"),_ 
						Browser("WebBrowser").Page("Cart").WebElement("ShoppingCartPanel").Link("ContinueShopping"),_ 
						Browser("WebBrowser").Page("Cart").WebElement("ShoppingCartPanel").WebButton("ProceedToCheckout"))
						
arrCheckResults = fnCheckPageElements(arrPageElements)

If arrCheckResults(0) Then
	fnReportStepEx "Pass", "Verify Cart subpage content.", "Cart subpage is displayed." & VbCrLf & "Current elements are displayed: " & arrCheckResults(1), Browser("WebBrowser"), "true"
Else 
	fnReportStepEx "Fail", "Verify Cart subpage content.", "Cart subpage is NOT displayed." & VbCrLf & " Current elements are NOT displayed: " & arrCheckResults(2), Browser("WebBrowser"), "true"
	ExitActionIteration "Verify_Cart.1"
End If

ExitActionIteration "0"

