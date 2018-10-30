'-----------------------------------------------------------------------------------------------------------------------------------------------
'Action Name : Navigate_HealthyLiving
'Description : This action is to verify that navigate to Healthy Living subpage is available
'Creation Date : 11.10.2018
'Last modification date : None
'Assumptions /Effects : Healthy Living subpage is displayed
'Returns : Action returns 0 if everything is correct or returns ActionNumber + step number if error occurs
'-----------------------------------------------------------------------------------------------------------------------------------------------
Option Explicit
Reporter.Filter = rfDisableAll

' Click on Healthy Living dropdown
If Browser("WebBrowser").Page("AllPages").WebElement("FitnessBlenderHeader").Link("HealthyLivingLink").Exist(5) Then
	Browser("WebBrowser").Page("AllPages").WebElement("FitnessBlenderHeader").Link("HealthyLivingLink").Click
Else
	fnReportStepEx "Fail", "Click on Healthy Living dropdown", "Healthy Living dropdown is NOT displayed", Browser("WebBrowser"), "true"
	ExitActionIteration "Navigate_HealthyLiving.1"
End If

'Verify Healthy Living subpage content
RunAction "Verify_HealthyLiving", oneIteration

ExitActionIteration "0"

