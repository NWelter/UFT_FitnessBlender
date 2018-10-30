'-----------------------------------------------------------------------------------------------------------------------------------------------
'Action Name : Verify_Fvorites
'Description : This action is to verify Favorites subpage content
'Creation Date : 24.10.2018
'Last modification date : None
'Assumptions /Effects : Favorites subpage content is correctly displayed
'Returns : Action returns 0 if everything is correct or returns ActionNumber + step number if error occurs
'-----------------------------------------------------------------------------------------------------------------------------------------------
Option Explicit
Reporter.Filter = rfDisableAll

Dim arrPageElements, arrCheckResults

'Verify Favorites subpage content
arrPageElements = Array(Browser("WebBrowser").Page("Favorites").WebElement("FilterSection"),_
						Browser("WebBrowser").Page("Favorites").WebElement("FavoriteVideosSection"),_ 
						Browser("WebBrowser").Page("Favorites").WebElement("FavoriteVideosSection").WebElement("VideoHeader"))

arrCheckResults = fnCheckPageElements(arrPageElements)

If arrCheckResults(0) Then
Browser("WebBrowser").Page("Favorites").WebElement("FavoriteVideosSection").Object.scrollIntoView
	fnReportStepEx "Pass", "Verify Favorites subpage content.", "Current elements are displayed: " & arrCheckResults(1), Browser("WebBrowser"), "true"
Else
	fnReportStepEx "Fail", "Verify Favorites subpage content.","Current elements are NOT displayed." & arrCheckResults(2), Browser("WebBrowser"), "true"
	ExitActionIteration "Verify_Favorites.1"
End If	

ExitActionIteration "0"

						
