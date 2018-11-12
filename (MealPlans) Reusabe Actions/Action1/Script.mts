'-----------------------------------------------------------------------------------------------------------------------------------------------
'Action Name : Verify_MealPlans
'Description : This action is to verify that Meal Plans subpage is displayed and verify content
'Author: Natalia Welter
'Creation Date : 11.10.2018
'Last modification date : None
'Assumptions /Effects : Meal Plans subpage and content are correctly displayed
'Returns : Action returns 0 if everything is correct or returns ActionNumber + step number if error occurs
'-----------------------------------------------------------------------------------------------------------------------------------------------
Option Explicit
Reporter.Filter = rfDisableAll

Dim arrPageElements, arrPageElementsRightDown, arrCheckResults 
Dim blnElementsAreDisplayed : blnElementsAreDisplayed = True

' Verify Meal Plans subpage main section content
arrPageElements = Array(Browser("WebBrowser").Page("MealPlans").WebElement("MealPlansHeader"),_ 
						Browser("WebBrowser").Page("MealPlans").WebElement("MainSection"),_
						Browser("WebBrowser").Page("MealPlans").WebElement("CalendarPlansHeader"),_ 
						Browser("WebBrowser").Page("MealPlans").WebElement("CalendarPlansSection"),_ 	
						Browser("WebBrowser").Page("MealPlans").WebElement("EbookPlansHeader"),_ 
						Browser("WebBrowser").Page("MealPlans").WebElement("EbookPlansSection"),_	
						Browser("WebBrowser").Page("MealPlans").WebElement("FeaturedHeader"),_ 
						Browser("WebBrowser").Page("MealPlans").WebElement("FeaturedSidebar"))

arrCheckResults = fnCheckPageElements(arrPageElements)

If arrCheckResults(0) Then
	fnReportStepEx "Pass", "Click on Meal Plans subtab. Verify main section content.",_ 
	"Meal Plans subpage is displayed." & VbCrLf & "Current elements of main section are displayed: " & arrCheckResults(1), Browser("WebBrowser"), "true"
Else 
	fnReportStepEx "Fail", "Click on Meal Plans subtab. Verify main section content.",_ 
	"Meal Plans subpage is NOT displayed." & VbCrLf & " Current elements of main section are NOT displayed: " & arrCheckResults(2), Browser("WebBrowser"), "true"
	blnElementsAreDisplayed = False
End If

' Verify Meal Plans subpage right sidebar content
arrPageElementsRightDown = Array (Browser("WebBrowser").Page("MealPlans").WebElement("FollowUsHeader"),_ 
						Browser("WebBrowser").Page("MealPlans").WebElement("SocialMediaContainer"),_ 
						Browser("WebBrowser").Page("MealPlans").WebElement("SocialMediaContainer").Link("Facebook"),_ 
						Browser("WebBrowser").Page("MealPlans").WebElement("SocialMediaContainer").Link("Google+"),_ 
						Browser("WebBrowser").Page("MealPlans").WebElement("SocialMediaContainer").Link("Instagram"),_ 
						Browser("WebBrowser").Page("MealPlans").WebElement("SocialMediaContainer").Link("Pinterest"),_ 
						Browser("WebBrowser").Page("MealPlans").WebElement("SocialMediaContainer").Link("Twitter"),_ 
						Browser("WebBrowser").Page("MealPlans").WebElement("SocialMediaContainer").Link("YouTube"))

arrCheckResults = fnCheckPageElements(arrPageElementsRightDown)

If arrCheckResults(0) Then
	Browser("WebBrowser").Page("MealPlans").WebElement("SocialMediaContainer").Object.scrollIntoView
	fnReportStepEx "Pass", "Click on Meal Plans subtab. Verify right sidebar content.",_ 
	"Meal Plans subpage is displayed." & VbCrLf & "Current elements of right sidebar are displayed: " & arrCheckResults(1), Browser("WebBrowser"), "true"
Else 
	Browser("WebBrowser").Page("MealPlans").WebElement("SocialMediaContainer").Object.scrollIntoView
	fnReportStepEx "Fail", "Click on Meal Plans subtab. Verify rigth sidebar content.",_ 
	"Meal Plans subpage is NOT displayed." & VbCrLf & " Current elements of right sidebar are NOT displayed: " & arrCheckResults(2), Browser("WebBrowser"), "true"
	blnElementsAreDisplayed = False
End If

If NOT blnElementsAreDisplayed Then
	ExitActionIteration "Verify_MealPlans.1"
End If

ExitActionIteration "0"
