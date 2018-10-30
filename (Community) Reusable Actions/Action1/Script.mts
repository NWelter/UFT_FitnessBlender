'-----------------------------------------------------------------------------------------------------------------------------------------------
'Action Name : Verify_Community
'Description : This action is to verify that Community subpage is displayed and verify content
'Creation Date : 11.10.2018
'Last modification date : None
'Assumptions /Effects : Community subpage and content are correctly displayed
'Returns : Action returns 0 if everything is correct or returns ActionNumber + step number if error occurs
'-----------------------------------------------------------------------------------------------------------------------------------------------
Option Explicit
Reporter.Filter = rfDisableAll

Dim arrPageElements, arrCheckResults

' Verify Community subpage content
arrPageElements = Array (Browser("WebBrowser").Page("Community").WebElement("CommunityHeader"),_ 
						Browser("WebBrowser").Page("Community").WebElement("DiscussionTopicsList"))
						
arrCheckResults = fnCheckPageElements(arrPageElements)

If arrCheckResults(0) Then
	fnReportStepEx "Pass", "Click on Community tab. Verify Community subpage content.", "Community subpage is displayed." & VbCrLf & "Current elements are displayed: " & arrCheckResults(1), Browser("WebBrowser"), "true"
Else 
	fnReportStepEx "Fail", "Click on Community tab. Verify Community subpage content.", "Community subpage is NOT displayed." & VbCrLf & " Current elements are NOT displayed: " & arrCheckResults(2), Browser("WebBrowser"), "true"
	ExitActionIteration "Verify_Community.1"
End If

ExitActionIteration "0"

