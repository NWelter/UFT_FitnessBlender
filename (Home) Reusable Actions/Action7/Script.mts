'------------------------------------------------------------------------------------------------------------
'Action Name: Redirect_Home
'Description: This action is to redirect to Home page
'Creation Date: 08-10-2018
'Last modification date: <None>
'Assumptions / Effects: Fitness Blender Home page is redirected and displayed correctly
'Returns: Action return 0 if everything is correct or returns ActionNumber + step number if error occure
'------------------------------------------------------------------------------------------------------------

Option Explicit
Reporter.Filter = rfDisableAll

' Click on Fitness Blender Logo button
If Browser("WebBrowser").Page("AllPages").WebElement("FitnessBlenderHeader").Link("LogoButton").Exist(30) Then
	Browser("WebBrowser").Page("AllPages").WebElement("FitnessBlenderHeader").Link("LogoButton").Click
	If fnIsVisible(Browser("WebBrowser").Page("Home").WebElement("MainContentSection")) Then
		fnReportStepEx "Pass", "Click on Fitness Blender Logo button", "Fitness Blender Logo button is displayed." & VbCrLf & "Fitness Blender Home page is redirected", Browser("WebBrowser"), "true"
	Else 
		fnReportStepEx "Fail", "Click on Fitness Blender Logo button", "Fitness Blender Logo button is displayed." & VbCrLf & "Fitness Blender Home page is NOT redirected", Browser("WebBrowser"), "true"
		ExitActionIteration "Redirect_Home.1.1"
	End  If
Else 
	fnReportStepEx "Fail", "Click on Fitness Blender Logo button", "Fitness Blender Logo button is NOT displayed." & VbCrLf & "Fitness Blender Home page is NOT redirected", Browser("WebBrowser"), "true"
	ExitActionIteration "Redirect_Home.1.2"
End If

ExitActionIteration "0"


	
