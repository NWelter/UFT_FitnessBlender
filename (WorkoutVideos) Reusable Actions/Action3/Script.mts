'Action Name : Select_VideoWithTitle
'Description : This action is to select specified workout video link with current title
'Creation Date : 24.10.2018
'Author: Natalia Welter
'Last modification date : None
'Inputs: strVideoTitle
'Assumptions /Effects : specified workout video subpage is displayed
'Returns : Action returns 0 if everything is correct or returns ActionNumber + step number if error occurs
'-----------------------------------------------------------------------------------------------------------------------------------------------
'Option Explicit
'Reporter.Filter = rfDisableAll
'
'Dim objVideoLinks, objObject, strSearchVideo, i
'
'Set objObject = Description.Create()
'	objObject("micclass").value = "Link"	
'Set objVideoLinks = Browser("WebBrowser").Page("WorkoutVideos").WebElement("WorkoutVideosSection").ChildObjects(objObject)
'
''Search specified video with title <<strVideoTitle>> and click on link
'For i = 0 To objVideoLinks.Count -1
'		strSearchVideo = Trim(objVideoLinks(i).GetROProperty("text"))
'			If strSearchVideo = Parameter("strVideoTitle") Then
'				objVideoLinks(i).Click
'				 fnReportStepEx "Pass", "Search specified video with title: " & Parameter("strVideoTitle") & " and click on link", "Video with title: " & Parameter("strVideoTitle") & "is found", Browser("WebBrowser"), "true"
'			Else
'				 fnReportStepEx "Fail", "Search specified video with title: " & Parameter("strVideoTitle") & " and click on link", "Video with title: " & Parameter("strVideoTitle") & "is NOT found", Browser("WebBrowser"), "true"
'				ExitActionIteration "Select_Video.1"				 
'		End If
'Next
'
'ExitActionIteration "0"
	

