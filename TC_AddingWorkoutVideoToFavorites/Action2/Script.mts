'------------------------------------------------------------------------------------------------------------
'Action Name: Recovery
'Description: This action is to return to starting state after test case
'Creation Date: 25-10-2018
'Author: Natalia Welter
'Last modification date: <None>
'Assumptions / Effects: Logged user doesn't have added videos on Favorites subpage
'Returns: Action return 0 if everything is correct or returns ActionNumber + step number if error occure
'------------------------------------------------------------------------------------------------------------
Option Explicit
Reporter.Filter = rfDisableAll
'
Dim objVideoLinks, objObject, strHeaderInfo, i

' Remove all added videos from Favorites subpage
Set objObject = Description.Create()
	objObject("micclass").value = "Link"	
Set objVideoLinks = Browser("WebBrowser").Page("Favorites").WebElement("FavoriteVideosSection").ChildObjects(objObject)

For i = 0 To objVideoLinks.Count -1
	Browser("WebBrowser").Page("Favorites").WebElement("FavoriteVideosSection").Link("FavoriteVideo").WebButton("HeartIcon").Click
	Browser("WebBrowser").Refresh
Next

strHeaderInfo = Trim(Browser("WebBrowser").Page("Favorites").WebElement("FavoriteVideosSection").WebElement("VideoHeader").GetROProperty("outertext"))

If strHeaderInfo = "No videos were found" Then
	fnReportStepEx"Pass", "Remove all added videos from Favorites subpage", "All videos are removed", Browser("WebBrowser"), "true"
Else
	fnReportStepEx "Fail", "Remove all added videos from Favorites subpage", "All videos are NOT removed", Browser("WebBrowser"), "true"
End If

ExitActionIteration "0"

