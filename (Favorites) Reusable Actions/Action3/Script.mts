'-----------------------------------------------------------------------------------------------------------------------------------------------
'Action Name : Check_VideoID
'Description : This action is to verify if proper video link is added to favorites
'Creation Date : 24.10.2018
'Last modification date : None
'Assumptions /Effects : Specified video link is added to favorites correctly
'Inputs: strVideoLinkID
'Returns : Action returns 0 if everything is correct or returns ActionNumber + step number if error occurs
'-----------------------------------------------------------------------------------------------------------------------------------------------
Option Explicit
Reporter.Filter = rfDisableAll

Dim strCurrentVideoID, objObject, objVideoLinks, i

Set objObject = Description.Create()
	objObject("micclass").value = "Link"
	Set objVideoLinks = Browser("WebBrowser").Page("Favorites").WebElement("FavoriteVideosSection").ChildObjects(objObject)

' Check if proper Video ID is added to favorites.
If Browser("WebBrowser").Page("Favorites").WebElement("FavoriteVideosSection").Link("FavoriteVideo").Exist(5) Then
	For i = 0 To objVideoLinks.Count -1  
		strCurrentVideoID = Trim(objVideoLinks(i).GetROProperty("html id"))
			If strCurrentVideoID = Parameter("strVideoLinkID") Then
				Browser("WebBrowser").Page("Favorites").WebElement("FavoriteVideosSection").Object.scrollIntoView
				fnReportStepEx "Pass", "Check if proper Video ID is added to favorites.",_ 
				"Specified Video is added to favorites. Current Video ID is: " & strCurrentVideoID , Browser("WebBrowser"), "true"
				Exit For	
			End If
	Next
		If strCurrentVideoId <> Parameter("strVideoLinkID") Then
				Browser("WebBrowser").Page("Favorites").WebElement("FavoriteVideosSection").Object.scrollIntoView
				fnReportStepEx "Fail", "Check if proper Video ID is added to favorites.",_ 
				"Specified Video is NOT added to favorites.", Browser("WebBrowser"), "true"
				ExitActionIteration "Check_VideoID.2"		
		End If
Else
	fnReportStepEx "Fail", "Check if proper Video ID is added to favorites.", "Video link is NOT display", Browser("WebBrowser"), "true"
	 ExitActionIteration "Check_VideoID.1"	
End If

ExitActionIteration "0"

