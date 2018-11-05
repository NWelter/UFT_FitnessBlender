'-----------------------------------------------------------------------------------------------------------------------------------------------
'Action Name : Add_To_Cart
'Description : This action is to add Meal Plan to Cart on Meal Plan details subpage
'Creation Date : 05.11.2018
'Author: Natalia Welter
'Last modification date : None
'Assumptions /Effects : Meal Plan is added to Cart
'Returns : Action returns 0 if everything is correct or returns ActionNumber + step number if error occurs
'-----------------------------------------------------------------------------------------------------------------------------------------------
Option Explicit
Reporter.Filter = rfDisableAll

'Click on Add To Bag button
If Browser("WebBrowser").Page("MealPlanDetails").WebElement("ProgramDetails").WebButton("AddToBag").Exist(5) Then
	Browser("WebBrowser").Page("MealPlanDetails").WebElement("ProgramDetails").WebButton("AddToBag").Click
Else
	fnReportStepEx "Fail", "Click on Add to Bag button", "Add to Bag button NOT exist", Browser("WebBrowser"), "true"
	ExitActionIteration "Add_To_Cart.1"
End I	f
'Verify redirection to Cart subpage
If Browser("WebBrowser").Page("Cart").Exist(15) Then
	fnReportStepEx "Pass", "Verify redirection to Cart subpage", "Cart subpage is displayed", Browser("WebBrowser"), "true"
Else
	fnReportStepEx "Fail", "Verify redirection to Cart subpage", "Cart subpage is NOT displayed", Browser("WebBrowser"), "true"
	ExitActionIteration "Add_To_Cart.2"
End If

'Verify Cart subpage content
RunAction "Verify_Cart [(Cart) Reusable Actions]", oneIteration

ExitActionIteration "0"
