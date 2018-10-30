'-----------------------------------------------------------------------------------------------------------------------------------------------
'Action Name : Navigate_BeforeAndAfter
'Description : This action is to verify that navigate to Before&After subpage is available
'Creation Date : 11.10.2018
'Last modification date : None
'Assumptions /Effects : Before&After subpage is displayed
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
    ExitActionIteration "Navigate_BeforeAndAfter.1"
End If

' Click on Before&After subtab
If Browser("WebBrowser").Page("AllPages").WebElement("FitnessBlenderHeader").WebElement("HealthyLivingDropdown").Link("Before&AfterSubtab").Exist(5) Then
	Browser("WebBrowser").Page("AllPages").WebElement("FitnessBlenderHeader").WebElement("HealthyLivingDropdown").Link("Before&AfterSubtab").Click
Else
	fnReportStepEx "Fail", "Click on Before&After subtab", "Before&After subtab is NOT displayed", Browser("WebBrowser"), "true"
	ExitActionIteration "Navigate_BeforeAndAfter.2"
End If

' Verify Before&After subpage content
RunAction "Verify_Before&After", oneIteration

ExitActionIteration "0"
