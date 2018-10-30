'-----------------------------------------------------------------------------------------------------------------------------------------------
'Action Name : Verify_HealthyRecipes
'Description : This action is to verify that Healthy Recipes subpage is displayed and verify content
'Creation Date : 10.10.2018
'Last modification date : None
'Assumptions /Effects : Healthy Recipes subpage and content are correctly displayed
'Returns : Action returns 0 if everything is correct or returns ActionNumber + step number if error occurs
'-----------------------------------------------------------------------------------------------------------------------------------------------
Option Explicit
Reporter.Filter = rfDisableAll

Dim arrPageElements, arrCheckResults

' Verify Healthy Recipes subpage content
arrPageElements = Array (Browser("WebBrowser").Page("HealthyRecipes").WebElement("HealthyRecipesHeader"),_ 
						Browser("WebBrowser").Page("HealthyRecipes").WebElement("HealthyRecipesSection"))

arrCheckResults = fnCheckPageElements(arrPageElements)

If arrCheckResults (0) Then
		fnReportStepEx "Pass", "Click on Healthy Recipes subtab. Verify Healthy Recipes subpage content.", "Healthy Recipes subpage is displayed." & VbCrLf &_ 
		"Current elements are displayed: " & arrCheckResults(1), Browser("WebBrowser"), "true"
Else
		fnReportStepEx "Fail", "Click on Healthy Recipes subtab. Verify Healthy Recipes subpage content.", "Healthy Recipes subpage is NOT displayed." & VbCrLf &_ 
		"Current elements are NOT displayed: " & arrCheckResults(2), Browser("WebBrowser"), "true"
		ExitActionIteration "Verify_HealthyRecipes.1"
End If

ExitActionIteration "0"

