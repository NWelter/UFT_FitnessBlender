'-----------------------------------------------------------------------------------------------------------------------------------------------
'Action Name : Verify_SignIn
'Description : This action is to verify that Sign In subpage is displayed and verify content
'Creation Date : 12.10.2018
'Last modification date : None
'Assumptions /Effects : Sign In subpage and content are correctly displayed
'Returns : Action returns 0 if everything is correct or returns ActionNumber + step number if error occurs
'-----------------------------------------------------------------------------------------------------------------------------------------------
Option Explicit
Reporter.Filter = rfDisableAll

Dim arrPageElements, arrCheckResults

' Verify Sign In subpage content
arrPageElements = Array (Browser("WebBrowser").Page("SignIn").Link("FacebookButton"),_
						Browser("WebBrowser").Page("SignIn").Link("Google+Button"),_
						Browser("WebBrowser").Page("SignIn").WebElement("SignInBox"),_ 
						Browser("WebBrowser").Page("SignIn").WebElement("SignInBox").WebEdit("Username"),_ 
						Browser("WebBrowser").Page("SignIn").WebElement("SignInBox").WebEdit("Password"),_ 
						Browser("WebBrowser").Page("SignIn").WebElement("SignInBox").WebElement("RememberMe"),_ 
						Browser("WebBrowser").Page("SignIn").WebElement("SignInBox").WebButton("SignInButton"))
						
arrCheckResults = fnCheckPageElements(arrPageElements)

If 	arrCheckResults(0) Then
	fnReportStepEx "Pass", "Verify Sign In subpage content",_ 
	"Sign In subpage is displayed." & VbCrLf & "Current elements are displayed: " & arrCheckResults(1), Browser("WebBrowser"), "true"
Else 
	fnReportStepEx "Fail", "Verify Sign In subpage content.",_ 
	"Sign In subpage is NOT displayed." & VbCrLf & "Current elements are displayed: " & arrCheckResults(2), Browser("WebBrowser"), "true"	
	ExitActionIteration "Verify_SignIn.1"
End If

ExitActionIteration "0"
