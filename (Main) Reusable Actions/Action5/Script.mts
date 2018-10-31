'-----------------------------------------------------------------------------------------------------------------------------------------------
'Action Name : Verify_Join
'Description : This action is to verify that Join subpage is displayed and verify content
'Creation Date : 12.10.2018
'Last modification date : None
'Assumptions /Effects : Join subpage and content are correctly displayed
'Returns : Action returns 0 if everything is correct or returns ActionNumber + step number if error occurs
'-----------------------------------------------------------------------------------------------------------------------------------------------
Option Explicit
Reporter.Filter = rfDisableAll

Dim arrPageElements, arrCheckResults

' Verify Join subpage content
arrPageElements = Array (Browser("WebBrowser").Page("Join").Link("FacebookButton"),_
						Browser("WebBrowser").Page("Join").Link("Google+Button"),_ 
						Browser("WebBrowser").Page("Join").WebElement("RegisterForm"),_
						Browser("WebBrowser").Page("Join").WebEdit("FirstName"),_ 
						Browser("WebBrowser").Page("Join").WebEdit("LastName"),_ 
						Browser("WebBrowser").Page("Join").WebEdit("Email"),_ 
						Browser("WebBrowser").Page("Join").WebEdit("Username"),_ 
						Browser("WebBrowser").Page("Join").WebEdit("Password"),_ 
						Browser("WebBrowser").Page("Join").WebEdit("ConfirmPassword"),_ 
						Browser("WebBrowser").Page("Join").WebButton("JoinButton"))
						
arrCheckResults = fnCheckPageElements(arrPageElements)

If arrCheckResults(0) Then
	fnReportStepEx "Pass", "Verify Join subpage content.",_ 
	"Join subpage is displayed." & VbCrLf & "Current elements are displayed: " & arrCheckResults(1), Browser("WebBrowser"), "true"
Else 
	fnReportStepEx "Fail", "Verify Join subpage content.",_ 
	"Join subpage is NOT displayed." & VbCrLf & "Current elements are displayed: " & arrCheckResults(2), Browser("WebBrowser"), "true"	
	ExitActionIteration "Verify_Join.1"
End If

ExitActionIteration "0"
