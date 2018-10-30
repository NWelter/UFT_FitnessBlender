'------------------------------------------------------------------------------------------------------------
'Action Name: Hover_MyFitness
'Description: This action is to verify that My Fitness dropdown is expanded and current subtabs are displayed
'Creation Date: 12-10-2018
'Last modification date: <None>
'Assumptions / Effects: My Fitness dropdown and current subtabs are displayed correctly
'Returns: Action return 0 if everything is correct or returns ActionNumber + step number if error occure
'------------------------------------------------------------------------------------------------------------

Option Explicit
Reporter.Filter = rfDisableAll

Dim arrPageElements, arrCheckResults

' Hover over My Fitness dropdown
If Browser("WebBrowser").Page("AllPages").WebElement("FitnessBlenderHeader").WebElement("MyFitness").Exist(5) Then
	Setting.WebPackage("ReplayType") = 2
	Browser("WebBrowser").Page("AllPages").WebElement("FitnessBlenderHeader").WebElement("MyFitness").FireEvent("onmouseover")
	Setting.WebPackage("ReplayType") = 1
	
	' Verify current buttons
	arrPageElements = Array(Browser("WebBrowser").Page("AllPages").WebElement("FitnessBlenderHeader").WebElement("MyFitnessMenu").Link("Join"),_ 
							Browser("WebBrowser").Page("AllPages").WebElement("FitnessBlenderHeader").WebElement("MyFitnessMenu").Link("SignIn"))

	arrCheckResults = fnCheckPageElements(arrPageElements)
	
	If arrCheckResults(0) Then
		fnReportStepEx "Pass", "Hover over My Fitness dropdown. Verify current buttons.", "My Fitness dropdown is displayed." & VbCrLf &_ 
		"Current buttons are displayed: " & arrCheckResults(1), Browser("WebBrowser"), "true"
	Else
		fnReportStepEx "Fail", "Hover over My Fitness dropdown. Verify current buttons.","My Fitness dropdown is displayed." & VbCrLf &_
		"Current buttons are NOT displayed." & arrCheckResults(2), Browser("WebBrowser"), "true"
		ExitActionIteration "Hover_MyFitness.1.1"
	End If		
Else
	fnReportStepEx "Fail", "Hover over My Fitness dropdown. Verify current buttons.", "My Fitness dropdown is NOT displayed", Browser("WebBrowser"), "true"
	ExitActionIteration "Hover_MyFitness.1.2"
End If 

ExitActionIteration "0"
