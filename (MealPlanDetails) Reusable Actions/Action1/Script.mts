'-----------------------------------------------------------------------------------------------------------------------------------------------
'Action Name : Verify_MealPlanDetails
'Description : This action is to verify Meal Plan details subpage content
'Creation Date : 02.11.2018
'Author: Natalia Welter
'Last modification date : None
'Assumptions /Effects : Meal Plan details subpage content is displayed correctly
'Returns : Action returns 0 if everything is correct or returns ActionNumber + step number if error occurs
'-----------------------------------------------------------------------------------------------------------------------------------------------
Option Explicit
Reporter.Filter = rfDisableAll

Dim arrPageElements, arrCheckResults

arrPageElements = Array(Browser("WebBrowser").Page("MealPlanDetails").WebElement("MealPlanHeader"),_ 
						Browser("WebBrowser").Page("MealPlanDetails").WebElement("MealPlanArticleBody"),_ 
						Browser("WebBrowser").Page("MealPlanDetails").WebElement("ProgramDetails"),_ 
						Browser("WebBrowser").Page("MealPlanDetails").WebElement("ProgramDetails").WebButton("AddToBag"))
						
arrCheckResults = fnCheckPageElements(arrPageElements)

' Verify Meal Plan details subpage content
If arrCheckResults(0) Then
	fnReportStepEx "Pass", "Verify Meal Plan details subpage content.",  "Current elements are displayed: " & arrCheckResults(1), Browser("WebBrowser"), "true"
Else 
	fnReportStepEx "Fail", "Verify Meal Plan details subpage content.", "Current elements are NOT displayed: " & arrCheckResults(2), Browser("WebBrowser"), "true"	
	ExitActionIteration "Verify_MealPlanDetails.1"
End If

ExitActionIteration "0"


