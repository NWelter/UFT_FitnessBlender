'-----------------------------------------------------------------------------------------------------------------------------------------------
'Action Name : Verify_Blog
'Description : This action is to verify that Blog subpage is displayed and verify content
'Creation Date : 11.10.2018
'Last modification date : None
'Assumptions /Effects : Blog subpage and content are correctly displayed
'Returns : Action returns 0 if everything is correct or returns ActionNumber + step number if error occurs
'-----------------------------------------------------------------------------------------------------------------------------------------------
Option Explicit
Reporter.Filter = rfDisableAll

Dim arrPageElements, arrPageElementsRightDown, arrCheckResults

' Verify Blog subpage main section content
arrPageElements = Array (Browser("WebBrowser").Page("Blog").WebElement("BlogHeader"),_ 
						Browser("WebBrowser").Page("Blog").WebElement("BlogArticlesSection"),_ 
						Browser("WebBrowser").Page("Blog").WebElement("FeaturedHeader"),_ 
						Browser("WebBrowser").Page("Blog").WebElement("FeaturedSidebar"))
						
arrCheckResults = fnCheckPageElements(arrPageElements)

If arrCheckResults(0) Then
	fnReportStepEx "Pass", "Click on Blog tab. Verify main section content.", "Blog subpage is displayed." & VbCrLf & "Current elements of main section are displayed: " & arrCheckResults(1), Browser("WebBrowser"), "true"
Else 
	fnReportStepEx "Fail", "Click on Blog tab. Verify main section content.", "Blog subpage is NOT displayed." & VbCrLf & " Current elements of main section are NOT displayed: " & arrCheckResults(2), Browser("WebBrowser"), "true"
	ExitActionIteration "Verify_Blog.1"
End If

' Verify Blog subpage right sidebar content
arrPageElementsRightDown = Array (Browser("WebBrowser").Page("Blog").WebElement("FollowUsHeader"),_
						Browser("WebBrowser").Page("Blog").WebElement("SocialMediaContainer"),_
						Browser("WebBrowser").Page("Blog").WebElement("SocialMediaContainer").Link("Facebook"),_ 
						Browser("WebBrowser").Page("Blog").WebElement("SocialMediaContainer").Link("Google+"),_ 
						Browser("WebBrowser").Page("Blog").WebElement("SocialMediaContainer").Link("Instagram"),_ 
						Browser("WebBrowser").Page("Blog").WebElement("SocialMediaContainer").Link("Pinterest"),_ 
						Browser("WebBrowser").Page("Blog").WebElement("SocialMediaContainer").Link("Twitter"),_ 
						Browser("WebBrowser").Page("Blog").WebElement("SocialMediaContainer").Link("Youtube"))

arrCheckResults = fnCheckPageElements(arrPageElementsRightDown)

If arrCheckResults(0) Then
	Browser("WebBrowser").Page("Blog").WebElement("SocialMediaContainer").Object.scrollIntoView
	fnReportStepEx "Pass", "Click on Blog tab. Verify right sidebar content.", "Blog subpage is displayed." & VbCrLf & "Current elements of the right sidebar are displayed: " & arrCheckResults(1), Browser("WebBrowser"), "true"
Else 
	fnReportStepEx "Fail", "Click on Blog tab. Verify right sidebar content.", "Blog subpage is NOT displayed." & VbCrLf & " Current elements of the right sidebar are NOT displayed: " & arrCheckResults(2), Browser("WebBrowser"), "true"
	ExitActionIteration "Verify_Blog.2"
End If

ExitActionIteration "0"

