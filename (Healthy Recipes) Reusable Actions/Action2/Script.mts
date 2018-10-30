'-----------------------------------------------------------------------------------------------------------------------------------------------
'Action Name : Navigate_HealthyRecipes
'Description : This action is to verify that navigate to Healthy Recipes subpage is available
'Creation Date : 11.10.2018
'Last modification date : None
'Assumptions /Effects : Healthy Recipes subpage is displayed
'Returns : Action returns 0 if everything is correct or returns ActionNumber + step number if error occurs
'-----------------------------------------------------------------------------------------------------------------------------------------------
Option Explicit
Reporter.Filter = rfDisableAll

Dim strRunActionStatus

' Hover over Healthy Living dropdown
strRunActionStatus = "9999"
strRunActionStatus = RunAction ("Hover_HealthyLivingDropdown [(Home) Reusable Actions]", oneIteration)
If (StrComp(strRunActionStatus, "0", 1) <> 0) Then
    fnReportStep "Fail", "Hover_HealthyLivingDropdown action failed" , "Returned value: " & strRunActionStatus , ""
    ExitActionIteration "Navigate_HealthyRecipes.1"
End If

' Click on Healthy Recipes subtab
If Browser("WebBrowser").Page("AllPages").WebElement("FitnessBlenderHeader").WebElement("HealthyLivingDropdown").Link("HealthyRecipesSubtab").Exist(5) Then
	Browser("WebBrowser").Page("AllPages").WebElement("FitnessBlenderHeader").WebElement("HealthyLivingDropdown").Link("HealthyRecipesSubtab").FireEvent("onclick")
Else
	fnReportStepEx "Fail", "Click on Healthy Recipes subtab", "Healthy Recipes subtab is NOT displayed", Browser("WebBrowser"), "true"
	ExitActionIteration "Navigate_HealthyRecipes.2"
End If

' Verify Healthy Recipes subpage content
RunAction "Verify_HealthyRecipes", oneIteration

ExitActionIteration "0"
