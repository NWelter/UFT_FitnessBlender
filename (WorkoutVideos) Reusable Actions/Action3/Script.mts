'Action Name : Select_VideoWithTitle
'Description : This action is to select specified workout video link with current title
'Creation Date : 24.10.2018
'Author: Natalia Welter
'Last modification date : None
'Inputs: strVideoTitle
'Assumptions /Effects : specified workout video subpage is displayed
'Returns : Action returns 0 if everything is correct or returns ActionNumber + step number if error occurs
'-----------------------------------------------------------------------------------------------------------------------------------------------
Option Explicit
Reporter.Filter = rfDisableAll

Dim strTitle : strTitle = ".*" & Parameter("strVideoTitle") & ".*"

'Select workout video link with title text <<strVideoTitle>>
Browser("WebBrowser").Page("WorkoutVideos").WebElement("WorkoutVideosSection").Link("VideoFirst").SetTOProperty "outertext", strTitle
If Browser("WebBrowser").Page("WorkoutVideos").WebElement("WorkoutVideosSection").Link("VideoFirst").Exist(5) Then
	Browser("WebBrowser").Page("WorkoutVideos").WebElement("WorkoutVideosSection").Link("VideoFirst").Click
Else
	fnReportStepEx "Fail", "Select video link with title text: " & "'" & Parameter("strVideoTitle") & "'",_ 
	"Workout video link with title text: " & "'" & Parameter("strVideoTitle") & "'" & " NOT exist", Browser("WebBrowser"), "true"
	ExitActionIteration "Select_VideoWithTitle.1"
End If

If Browser("WebBrowser").Page("WorkoutVideosDetails").WebElement("WorkoutDetailsPanel").Exist(10) Then
	fnReportStepEx "Pass", "Select video link with title text: " & "'" & Parameter("strVideoTitle") & "'",_ 
	"Workout video link with title text: " & "'" & Parameter("strVideoTitle") & "'" & " is selected", Browser("WebBrowser"), "true"	
Else
	fnReportStepEx "Fail", "Select video link with title text: " & "'" & Parameter("strVideoTitle") & "'",_ 
	"Workout video link with title text: " & "'" & Parameter("strVideoTitle") & "'" & " is NOT selected", Browser("WebBrowser"), "true"
	ExitActionIteration "Select_VideoWithTitle.2"	
End If

ExitActionIteration "0"
	

