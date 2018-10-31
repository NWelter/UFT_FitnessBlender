'-----------------------------------------------------------------------------------------------------------------------------------------------
'Action Name : Verify_NewAccount
'Description : This action is to verify that New Account subpage is displayed and verify content
'Creation Date : 29.10.2018
'Last modification date : None
'Assumptions /Effects : New Account subpage and content are correctly displayed
'Returns : Action returns 0 if everything is correct or returns ActionNumber + step number if error occurs
'-----------------------------------------------------------------------------------------------------------------------------------------------
Option Explicit
Reporter.Filter = rfDisableAll

Dim arrPageElements, arrCheckResults

' Verify New Account subpage content
arrPageElements = Array(Browser("WebBrowser").Page("NewAccount").WebElement("MemberProfileForm"),_ 
						Browser("WebBrowser").Page("NewAccount").WebElement("MemberProfileForm").WebElement("ProfileFormHeader"),_ 
						Browser("WebBrowser").Page("NewAccount").WebElement("MemberProfileForm").WebEdit("FirstName"),_ 
						Browser("WebBrowser").Page("NewAccount").WebElement("MemberProfileForm").WebEdit("LastName"),_ 
						Browser("WebBrowser").Page("NewAccount").WebElement("MemberProfileForm").WebEdit("DisplayName"),_ 
						Browser("WebBrowser").Page("NewAccount").WebElement("MemberProfileForm").WebEdit("Email"),_ 
						Browser("WebBrowser").Page("NewAccount").WebElement("MemberProfileForm").WebElement("ProfileImageUpload"),_
						Browser("WebBrowser").Page("NewAccount").WebCheckBox("Subscribe"),_ 
						Browser("WebBrowser").Page("NewAccount").WebButton("CreateAccount"))
						
arrCheckResults = fnCheckPageElements(arrPageElements)

If 	arrCheckResults(0) Then
	fnReportStepEx "Pass", "Verify New Account subpage content",_ 
	"New Account subpage is displayed." & VbCrLf & "Current elements are displayed: " & arrCheckResults(1), Browser("WebBrowser"), "true"
Else 
	fnReportStepEx "Fail", "Verify New Account subpage content.",_ 
	"New Account subpage is NOT displayed." & VbCrLf & "Current elements are displayed: " & arrCheckResults(2), Browser("WebBrowser"), "true"	
	ExitActionIteration "Verify_NewAccount.1"
End If

ExitActionIteration "0"
