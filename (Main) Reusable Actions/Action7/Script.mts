'-----------------------------------------------------------------------------------------------------------------------------------------------
'Action Name : Navigate_Join
'Description : This action is to verify that navigate to Join subpage is available
'Creation Date : 11.10.2018
'Last modification date : None
'Assumptions /Effects : Join subpage is displayed
'Returns : Action returns 0 if everything is correct or returns ActionNumber + step number if error occurs
'-----------------------------------------------------------------------------------------------------------------------------------------------
Option Explicit
Reporter.Filter = rfDisableAll

Dim strRunActionStatus

' Hover over My Fitness dropdown
strRunActionStatus = "9999"
strRunActionStatus = RunAction ("Hover_MyFitness [(Home) Reusable Actions]", oneIteration)
If (StrComp(strRunActionStatus, "0", 1) <> 0) Then
    fnReportStep "Fail", "Hover_MyFitness action failed" , "Returned value: " & strRunActionStatus , ""
    ExitActionIteration "Navigate_Join.1"
End If

' Click on Join button
If Browser("WebBrowser").Page("AllPages").WebElement("FitnessBlenderHeader").WebElement("MyFitnessMenu").Link("Join").Exist(5) Then
	Browser("WebBrowser").Page("AllPages").WebElement("FitnessBlenderHeader").WebElement("MyFitnessMenu").Link("Join").Click
Else
	fnReportStepEx "Fail", "Click on Join button", "Join button is NOT displayed", Browser("WebBrowser"), "true"
	ExitActionIteration "Navigate_Join.2"	
End If

'Verify Join subpage content
RunAction "Verify_Join", oneIteration

ExitActionIteration "0"
