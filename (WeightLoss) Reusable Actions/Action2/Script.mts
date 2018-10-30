'-----------------------------------------------------------------------------------------------------------------------------------------------
'Action Name : Navigate_WeightLoss
'Description : This action is to verify that navigate to Weight Loss subpage is available
'Creation Date : 11.10.2018
'Last modification date : None
'Assumptions /Effects : Weight Loss subpage is displayed
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
    ExitActionIteration "Navigate_WeightLoss.1"
End If

' Click on Weight Loss subtab
If Browser("WebBrowser").Page("AllPages").WebElement("FitnessBlenderHeader").WebElement("HealthyLivingDropdown").Link("WeightLossSubtab").Exist(5) Then
	Browser("WebBrowser").Page("AllPages").WebElement("FitnessBlenderHeader").WebElement("HealthyLivingDropdown").Link("WeightLossSubtab").Click
Else
	fnReportStepEx "Fail", "Click on Weight Loss subtab", "Weight Loss subtab is NOT displayed", Browser("WebBrowser"), "true"
	ExitActionIteration "Navigate_WeightLoss.2"
End If

' Verify Weigth Loss subpage content
RunAction "Verify_WeightLoss", oneIteration

ExitActionIteration "0"
