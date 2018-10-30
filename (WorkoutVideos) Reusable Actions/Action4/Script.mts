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

'Click on first Workout Video link
If Browser("WebBrowser").Page("WorkoutVideos").WebElement("WorkoutVideosSection").Link("VideoFirst").Exist(5) Then
	
	'Get title from video link 
	strVideoLinkTitle = Trim(Browser("WebBrowser").Page("WorkoutVideos").WebElement("WorkoutVideosSection").Link("VideoFirst").WebElement("VideoTitle").GetROProperty("outertext"))
	'Get ID from video link to <<strVideoLinkID>>
	Parameter ("strVideoLinkID") = Trim(Browser("WebBrowser").Page("WorkoutVideos").WebElement("WorkoutVideosSection").Link("VideoFirst").GetROProperty("html id"))	
	
	Browser("WebBrowser").Page("WorkoutVideos").WebElement("WorkoutVideosSection").Object.scrollIntoView
	Browser("WebBrowser").Page("WorkoutVideos").WebElement("WorkoutVideosSection").Link("VideoFirst").Click
	fnReportStepEx "Pass", "Click on first workout video link", "Workout video details subpage is displayed", Browser("WebBrowser"), "true"
Else
	fnReportStepEx "Fail", "Click on first workout video link", "Workout video details subpage is NOT displayed", Browser("WebBrowser"), "true"
	ExitActionIteration "Select_Video.1"
End If

'Verify Workout Video details subpage content
strRunActionStatus = "9999"
strRunActionStatus = RunAction ("Verify_WorkoutVideosDetails [(WorkoutVideosDetails)Reusable Actions]", oneIteration)
If (StrComp(strRunActionStatus, "0", 1) <> 0) Then
    fnReportStep "Fail", "Verify_WorkoutVideosDetails action failed" , "Returned value: " & strRunActionStatus , ""
    ExitActionIteration "Select_Video.2"
End If

'Check if select link title is equal to video details header
If Browser("WebBrowser").Page("WorkoutVideosDetails").WebElement("VideoHeader").Exist(5) Then	
	strVideoHeader = Trim(Browser("WebBrowser").Page("WorkoutVideosDetails").WebElement("VideoHeader").GetROProperty("outertext"))
		If strVideoHeader <> strVideoLinkTitle Then
			fnReportStepEx "Fail", "Check if select link title is equal to video details header.", "Workout Video details header is NOT equal to select link title", Browser("WebBrowser"), "true"
			ExitActionIteration "Select_Video.3.1"
		End If
Else
	fnReportStepEx "Fail", "Check if select link title is equal to video details header", "Workout Video details header is NOT displayed", Browser("WebBrowser"), "true"
	ExitActionIteration "Select_Video.3.2"
End  If

ExitActionIteration "0"
