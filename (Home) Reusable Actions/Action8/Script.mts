'-----------------------------------------------------------------------------------------------------------------------------------------------
'Action Name : Verify_FitnessBlenderHeader
'Description : This action is to verify that Fitness Blender navbar header is displayed and verify content
'Creation Date : 05.10.2018
'Last modification date : None
'Assumptions /Effects : Fitness Blender navbar header elements are correctly displayed
'Returns : Action returns 0 if everything is correct or returns ActionNumber + step number if error occurs
'-----------------------------------------------------------------------------------------------------------------------------------------------
Option Explicit
Reporter.Filter = rfDisableAll

Dim arrPageElements, arrCheckResults

'Verify elements on Fitness Blender navbar header
arrPageElements = Array(Browser("WebBrowser").Page("AllPages").WebElement("FitnessBlenderHeader"),_
				Browser("WebBrowser").Page("AllPages").WebElement("FitnessBlenderHeader").Link("LogoButton"),_
				Browser("WebBrowser").Page("AllPages").WebElement("FitnessBlenderHeader").WebElement("Workouts&ProgramsDropdown"),_ 
				Browser("WebBrowser").Page("AllPages").WebElement("FitnessBlenderHeader").WebElement("HealthyLivingDropdown"),_
				Browser("WebBrowser").Page("AllPages").WebElement("FitnessBlenderHeader").WebElement("Community"),_
				Browser("WebBrowser").Page("AllPages").WebElement("FitnessBlenderHeader").WebElement("Blog"),_
				Browser("WebBrowser").Page("AllPages").WebElement("FitnessBlenderHeader").WebElement("About"),_
				Browser("WebBrowser").Page("AllPages").WebElement("FitnessBlenderHeader").WebElement("MyFitness"),_
				Browser("WebBrowser").Page("AllPages").WebElement("FitnessBlenderHeader").WebElement("ScopeIcon"),_
				Browser("WebBrowser").Page("AllPages").WebElement("FitnessBlenderHeader").WebElement("ShoppingbagIcon"))
				
arrCheckResults = fnCheckPageElements(arrPageElements)

If arrCheckResults(0) Then
	fnReportStepEx "Pass", "Verify elements on Fitness Blender navbar header", "Current elements on Fitness Blender navbar header are displayed: " & arrCheckResults(1), Browser("WebBrowser"), "true"
Else
	fnReportStepEx "Fail", "Verify elements on Fitness Blender navbar header", "Current elements on Fitness Blender navbar header are NOT displayed: " & arrCheckResults(2), Browser("WebBrowser"), "true"
	ExitActionIteration "Verify_FitnessBlenderHeader.1"	
End If

ExitActionIteration "0"
