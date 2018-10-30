'-----------------------------------------------------------------------------------------------------------------------------------------------
'Action Name : Navigate_Health
'Description : This action is to verify that navigate to Health subpage is available
'Creation Date : 11.10.2018
'Last modification date : None
'Assumptions /Effects : Health subpage is displayed
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
    ExitActionIteration "Navigate_Health.1"
End If

' Click on Health subtab
If Browser("WebBrowser").Page("AllPages").WebElement("FitnessBlenderHeader").WebElement("HealthyLivingDropdown").Link("HealthSubtab").Exist(5) Then
	Browser("WebBrowser").Page("AllPages").WebElement("FitnessBlenderHeader").WebElement("HealthyLivingDropdown").Link("HealthSubtab").Click
Else
	fnReportStepEx "Fail", "Click on Health subtab", "Health subtab is NOT displayed", Browser("WebBrowser"), "true"
	ExitActionIteration "Navigate_Health.2"
End If

' Verify Health subpage content
RunAction "Verify_Health", oneIteration

ExitActionIteration "0"

