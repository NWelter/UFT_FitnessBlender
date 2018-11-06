'-----------------------------------------------------------------------------------------------------------------------------------------------
'Action Name : Select_Video
'Description : This action is to select Workout Video link on Workout Videos subpage
'Creation Date : 24.10.2018
'Author: Natalia Welter
'Last modification date : None
'Outputs: strVideoLinkID
'Assumptions /Effects : Workout Video details subpage is displayed
'Returns : Action returns 0 if everything is correct or returns ActionNumber + step number if error occurs
'-----------------------------------------------------------------------------------------------------------------------------------------------
Option Explicit
Reporter.Filter = rfDisableAll

Dim strRunActionStatus, strVideoLinkTitle, strVideoHeader


'Click on Workout Video link
If Browser("WebBrowser").Page("WorkoutVideos").WebElement("WorkoutVideosSection").Link("Video").Exist(5) Then
	
	'Get title from video link 
	strVideoLinkTitle = Trim(Browser("WebBrowser").Page("WorkoutVideos").WebElement("WorkoutVideosSection").Link("Video").WebElement("VideoTitle").GetROProperty("outertext"))
	
	'Get ID from video link to <<strVideoLinkID>>
	Parameter ("strVideoLinkID") = Trim(Browser("WebBrowser").Page("WorkoutVideos").WebElement("WorkoutVideosSection").Link("Video").GetROProperty("html id"))	
	
	Browser("WebBrowser").Page("WorkoutVideos").WebElement("WorkoutVideosSection").Link("Video").Click
Else
	Browser("WebBrowser").Page("WorkoutVideos").WebElement("WorkoutVideosSection").Object.scrollIntoView
	fnReportStepEx "Fail", "Click on workout video link", "Workout video link NOT exist", Browser("WebBrowser"), "true"
	ExitActionIteration "Select_Video.1"
End If

If Browser("WebBrowser").Page("WorkoutVideosDetails").WebElement("WorkoutDetailsPanel").Exist(10) Then
	fnReportStepEx "Pass", "Click on workout video link", "Workout video details subpage is displayed", Browser("WebBrowser"), "true"
Else 
	fnReportStepEx "Fail", "Click on workout video link", "Workout video details subpage is NOT displayed", Browser("WebBrowser"), "true"
	ExitActionIteration "Select_Video.2"
End If
	
'Verify Workout Video details subpage content
RunAction "Verify_WorkoutVideosDetails [(WorkoutVideosDetails)Reusable Actions]", oneIteration

'Check if select link title is equal to video details header
If Browser("WebBrowser").Page("WorkoutVideosDetails").WebElement("VideoHeader").Exist(5) Then	
	strVideoHeader = Trim(Browser("WebBrowser").Page("WorkoutVideosDetails").WebElement("VideoHeader").GetROProperty("outertext"))
Else
	fnReportStepEx "Fail", "Check if select link title is equal to video details header.",_ 
	"Workout Video details header is NOT displayed", Browser("WebBrowser"), "true"
	ExitActionIteration "Select_Video.3"
End  If

If strVideoHeader = strVideoLinkTitle Then
	fnReportStepEx "Pass", "Check if select link title is equal to video details header.",_ 
	"Workout Video details header is equal to select link title: "& VbCrLf & "'" & strVideoLinkTitle & "'", Browser("WebBrowser"), "true"
Else
	fnReportStepEx "Fail", "Check if select link title is equal to video details header.",_ 
	"Workout Video details header is NOT equal to select link title: "& VbCrLf & "'" & strVideoLinkTitle & "'", Browser("WebBrowser"), "true"
	ExitActionIteration "Select_Video.4"
End If

ExitActionIteration "0"
