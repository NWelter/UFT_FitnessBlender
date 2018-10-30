'Action Name : Navigate_Fitness
'Description : This action is to verify that navigate to Fitness subpage is available
'Creation Date : 11.10.2018
'Last modification date : None
'Assumptions /Effects : Fitness subpage is displayed
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
    ExitActionIteration "Navigate_Fitness.1"
End If

' Click on Fitness subtab
If Browser("WebBrowser").Page("AllPages").WebElement("FitnessBlenderHeader").WebElement("HealthyLivingDropdown").Link("FitnessSubtab").Exist(5) Then
	Browser("WebBrowser").Page("AllPages").WebElement("FitnessBlenderHeader").WebElement("HealthyLivingDropdown").Link("FitnessSubtab").Click
Else
	fnReportStepEx "Fail", "Click on Fitness subtab", "Fitness subtab is NOT displayed", Browser("WebBrowser"), "true"
	ExitActionIteration "Navigate_Fitness.2"
End If

' Verify Fitness subpage content
RunAction "Verify_Fitness", oneIteration

ExitActionIteration "0"

