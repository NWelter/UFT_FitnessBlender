'-----------------------------------------------------------------------------------------------------------------------------------------------
'Action Name : Compare_ItemsAmount
'Description : This action is to compare if valid amount of added items to Cart is displayed in Shopping Bag icon
'Creation Date : 05.11.2018
'Last modification date : None
'Assumptions /Effects : Amount of added items is displayed in Shopping Bag icon correctly
'Returns : Action returns 0 if everything is correct or returns ActionNumber + step number if error occurs
'-----------------------------------------------------------------------------------------------------------------------------------------------
Option Explicit
Reporter.Filter = rfDisableAll

Dim intAmountInBagIcon, intNumberOfItemsInCart, objObject, objItemsLinks

' Get items amount displayed in Shopping Bag icon on navbar header
If Browser("WebBrowser").Page("AllPages").WebElement("FitnessBlenderHeader").WebElement("CartItemsNumber").Exist(5) Then
	intAmountInBagIcon = CInt(Browser("WebBrowser").Page("AllPages").WebElement("FitnessBlenderHeader").WebElement("CartItemsNumber").GetROProperty("outertext"))
Else
	fnReportStepEx "Fail", "Check items amount in Shopping Bag icon on navbar header", "Items amount in Shopping Bag icon is NOT displayed.", Browser("WebBrowser"), "true"
	ExitActionIteration "Compare_ItemsAmount.1"
End If

' Get number of items link added to Cart
If Browser("WebBrowser").Page("Cart").WebElement("ShoppingCartPanel").Link("ItemLink").Exist(5) Then
	Set objObject = Description.Create()
		objObject("micclass").value = "Link"
	Set objItemsLinks = Browser("WebBrowser").Page("Cart").WebElement("ShoppingCartPanel").ChildObjects(objObject)
	intNumberOfItemsInCart = objItemsLinks.Count -1
Else
	fnReportStepEx "Fail", "Get number of items links added to Cart", "Items links are NOT displayed.", Browser("WebBrowser"), "true"
	ExitActionIteration "Compare_ItemsAmount.2"
End If

' Compare amount displayed in Shopping Bag icon with number of items links added to Cart
If intAmountInBagIcon = intNumberOfItemsInCart Then
	Browser("WebBrowser").Page("AllPages").WebElement("FitnessBlenderHeader").WebElement("CartItemsNumber").Object.scrollIntoView
	fnReportStepEx "Pass", "Compare amount displayed in Shopping Bag icon with number of items links added to Cart.",_ 
	"Amount of items is displayed in Shopping Bag icon correctly.", Browser("WebBrowser"), "true"
Else
	Browser("WebBrowser").Page("AllPages").WebElement("FitnessBlenderHeader").WebElement("CartItemsNumber").Object.scrollIntoView
	fnReportStepEx "Fail", "Compare amount displayed in Shopping Bag icon with number of items links added to Cart.",_ 
	"Amount of items is displayed in Shopping Bag icon NOT correctly." & VbCrLf & _ 
	"Number of items in Cart: " & intNumberOfItemsInCart & VbCrLf & _ 
	"Amount in Shopping Bag icon: " & intAmountInBagIcon, Browser("WebBrowser"), "true"
	ExitActionIteration "Compare_ItemsAmount.3"
End If

ExitActionIteration "0"
 
