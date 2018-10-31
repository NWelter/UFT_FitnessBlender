'------------------------------------------------------------------------------------------------------------
'Action Name: Hover_HealthyLivingDropdown
'Description: This action is to verify that Healthy Living dropdown is expanded and current subtabs are displayed
'Creation Date: 08-10-2018
'Last modification date: <None>
'Assumptions / Effects: Healthy Living dropdown and current subtabs are displayed correctly
'Returns: Action return 0 if everything is correct or returns ActionNumber + step number if error occure
'------------------------------------------------------------------------------------------------------------

Option Explicit
Reporter.Filter = rfDisableAll

Dim arrPageElements, arrCheckResults

' Hover over Healthy Living dropdown
If Browser("WebBrowser").Page("AllPages").WebElement("FitnessBlenderHeader").WebElement("HealthyLivingDropdown").Exist(5) Then
	Setting.WebPackage("ReplayType") = 2
	Browser("WebBrowser").Page("AllPages").WebElement("FitnessBlenderHeader").WebElement("HealthyLivingDropdown").FireEvent("onmouseover")
	Setting.WebPackage("ReplayType") = 1
	
	' Verify current subtabs	
	arrPageElements = Array(Browser("WebBrowser").Page("AllPages").WebElement("FitnessBlenderHeader").WebElement("HealthyLivingDropdown").Link("FitnessSubtab"),_ 
							Browser("WebBrowser").Page("AllPages").WebElement("FitnessBlenderHeader").WebElement("HealthyLivingDropdown").Link("HealthSubtab"),_ 
							Browser("WebBrowser").Page("AllPages").WebElement("FitnessBlenderHeader").WebElement("HealthyLivingDropdown").Link("HealthyRecipesSubtab"),_ 
							Browser("WebBrowser").Page("AllPages").WebElement("FitnessBlenderHeader").WebElement("HealthyLivingDropdown").Link("Before&AfterSubtab"),_ 
							Browser("WebBrowser").Page("AllPages").WebElement("FitnessBlenderHeader").WebElement("HealthyLivingDropdown").Link("WeightLossSubtab"))
	
	arrCheckResults = fnCheckPageElements(arrPageElements)
	
	If arrCheckResults (0) Then
		fnReportStepEx "Pass", "Hover over Healthy Living dropdown. Verify current subtabs.", "Healthy Living dropdown is displayed." & VbCrLf &_ 
		"Current subtabs are displayed: " & arrCheckResults(1), Browser("WebBrowser"), "true"
	Else
		fnReportStepEx "Fail", "Hover over Healthy Living dropdown. Verify current subtabs.", "Healthy Living dropdown is displayed." & VbCrLf &_ 
		"Current subtabs are NOT displayed: " & arrCheckResults(2), Browser("WebBrowser"), "true"
		ExitActionIteration "Hover_HealthyLivingDropdown.2"
	End If	
Else 
	fnReportStepEx "Fail", "Hover over Healthy Living dropdown. Verify current subtabs.", "Healthy Living dropdown is NOT displayed", Browser("WebBrowser"), "true"
	ExitActionIteration "Hover_HealthyLivingDropdown.1"
End If

ExitActionIteration "0"
