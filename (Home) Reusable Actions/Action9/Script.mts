'------------------------------------------------------------------------------------------------------------
'Action Name: Hover_MyFitnessLoggedUser
'Description: This action is to verify that My Fitness dropdown for logged user is expanded and current subtabs are displayed
'Creation Date: 12-10-2018
'Last modification date: <None>
'Assumptions / Effects: My Fitness dropdown for logged user and current subtabs are displayed correctly
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
	
	' Verify current buttons for logged user
	arrPageElements = Array(Browser("WebBrowser").Page("AllPages").WebElement("FitnessBlenderHeader").WebElement("MyFitnessMenuUser").Link("Dashboard"),_ 
							Browser("WebBrowser").Page("AllPages").WebElement("FitnessBlenderHeader").WebElement("MyFitnessMenuUser").Link("Calendar"),_ 
							Browser("WebBrowser").Page("AllPages").WebElement("FitnessBlenderHeader").WebElement("MyFitnessMenuUser").Link("PurchasedPrograms"),_ 
							Browser("WebBrowser").Page("AllPages").WebElement("FitnessBlenderHeader").WebElement("MyFitnessMenuUser").Link("Favorites"),_ 
							Browser("WebBrowser").Page("AllPages").WebElement("FitnessBlenderHeader").WebElement("MyFitnessMenuUser").Link("Notifications"),_ 
							Browser("WebBrowser").Page("AllPages").WebElement("FitnessBlenderHeader").WebElement("MyFitnessMenuUser").Link("Account"),_ 
							Browser("WebBrowser").Page("AllPages").WebElement("FitnessBlenderHeader").WebElement("MyFitnessMenuUser").Link("SignOut"))
							
	arrCheckResults = fnCheckPageElements(arrPageElements)
	
	If arrCheckResults(0) Then
		fnReportStepEx "Pass", "Hover over My Fitness dropdown. Verify current buttons for logged user.", "My Fitness dropdown is displayed." & VbCrLf &_ 
		"Current buttons for logged user are displayed: " & arrCheckResults(1), Browser("WebBrowser"), "true"
	Else
		fnReportStepEx "Fail", "Hover over My Fitness dropdown. Verify current buttons for logged user.","My Fitness dropdown is displayed." & VbCrLf &_
		"Current buttons for logged user are NOT displayed." & arrCheckResults(2), Browser("WebBrowser"), "true"
		ExitActionIteration "Hover_MyFitnessLoggedUser.1.1"
	End If		
Else
	fnReportStepEx "Fail", "Hover over My Fitness dropdown. Verify current buttons for logged user.", "My Fitness dropdown is NOT displayed", Browser("WebBrowser"), "true"
	ExitActionIteration "Hover_MyFitnessLoggedUser.1.2"
End If 

ExitActionIteration "0"
