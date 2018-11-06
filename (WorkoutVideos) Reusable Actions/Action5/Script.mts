'-----------------------------------------------------------------------------------------------------------------------------------------------
'Action Name : Add_To_Favorites_By_Icon
'Description : This action is to add video to favorites on Workout Videos subpage by Heart icon
'Creation Date : 24.10.2018
'Author: Natalia Welter
'Last modification date : None
'Inputs: strIconColor, strBorderColor
'Outputs: strVideoLinkID
'Assumptions /Effects : Workout Video is added to favorites
'Returns : Action returns 0 if everything is correct or returns ActionNumber + step number if error occurs
'-----------------------------------------------------------------------------------------------------------------------------------------------
Option Explicit
Reporter.Filter = rfDisableAll

Dim strCurrentIconColor, strCurrentBorderColor

'Select Workout Video by index:
Browser("WebBrowser").Page("WorkoutVideos").WebElement("WorkoutVideosSection").Link("Video").SetTOProperty "index", Parameter ("intVideoLinkOrder") -1

' Hover over Workout Video from the list
If Browser("WebBrowser").Page("WorkoutVideos").WebElement("WorkoutVideosSection").Link("Video").Exist(5) Then
	Setting.WebPackage("ReplayType") = 2
	Browser("WebBrowser").Page("WorkoutVideos").WebElement("WorkoutVideosSection").Link("Video").FireEvent("onmouseover")
	
	'Get ID from video link to <<strVideoLinkID>>
	Parameter ("strVideoLinkID") = Trim(Browser("WebBrowser").Page("WorkoutVideos").WebElement("WorkoutVideosSection").Link("Video").GetROProperty("html id"))
Else
	Browser("WebBrowser").Page("WorkoutVideos").WebElement("WorkoutVideosSection").Link("Video").Object.scrollIntoView
	fnReportStepEx "Fail", "Hover over Workout Video from the list and click on Heart icon.", "Workout Video is NOT displayed", Browser("WebBrowser"), "true"
	ExitActionIteration "Add_T0_Favorites_By_Icon.1"	
End If

' Verify border color of selected video link
strCurrentBorderColor = Browser("WebBrowser").Page("WorkoutVideos").WebElement("WorkoutVideosSection").Link("Video").GetROProperty("style/border-color")
If strCurrentBorderColor = Parameter ("strBorderColor") Then
	fnReportStepEx "Pass", "Verify border color of selected video link.", "Border color is equal to: " & Parameter ("strBorderColor"), Browser("WebBrowser"), "true"
Else
	fnReportStepEx "Fail", "Verify border color of selected video link.",_ 
	"Border color is NOT equal to: " & Parameter ("strBorderColor") & VbCrLf & "Current border color is: " & strCurrentBorderColor, Browser("WebBrowser"), "true"
End If

' Click on Heart icon
If Browser("WebBrowser").Page("WorkoutVideos").WebElement("WorkoutVideosSection").Link("Video").WebButton("HeartIcon").Exist(5) Then
	Browser("WebBrowser").Page("WorkoutVideos").WebElement("WorkoutVideosSection").Link("Video").WebButton("HeartIcon").Click
	Setting.WebPackage("ReplayType") = 1
	Browser("WebBrowser").Page("WorkoutVideos").WebElement("WorkoutVideosSection").Link("Video").Object.scrollIntoView
	fnReportStepEx "Pass", "Click on Heart icon.", "Heart icon is displayed.", Browser("WebBrowser"), "true"	
Else
	Browser("WebBrowser").Page("WorkoutVideos").WebElement("WorkoutVideosSection").Link("Video").Object.scrollIntoView
	fnReportStepEx "Fail", "Click on Heart icon.", "Heart icon NOT exist.", Browser("WebBrowser"), "true"
	ExitActionIteration "Add_To_Favorites_By_Icon.2"		
End If

' Verify color of selected Heart icon
strCurrentIconColor = Browser("WebBrowser").Page("WorkoutVideos").WebElement("WorkoutVideosSection").Link("Video").WebButton("HeartIcon").GetROProperty("style/color")
If strCurrentIconColor = Parameter ("strIconColor") Then
	Browser("WebBrowser").Page("WorkoutVideos").WebElement("WorkoutVideosSection").Link("Video").Object.scrollIntoView
	fnReportStepEx "Pass", "Verify color of selected Heart icon", "Heart icon color is equal to: " & Parameter ("strIconColor"), Browser("WebBrowser"), "true"
Else 
	Browser("WebBrowser").Page("WorkoutVideos").WebElement("WorkoutVideosSection").Link("Video").Object.scrollIntoView
	fnReportStepEx "Fail", "Verify color of selected Heart icon",_ 
	"Heart icon color is NOT equal to: " & Parameter ("strIconColor") & VbCrLf & "Current icon color is: " & strCurrentIconColor, Browser("WebBrowser"), "true"
End If

ExitActionIteration "0"
