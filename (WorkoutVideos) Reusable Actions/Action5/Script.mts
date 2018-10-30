'-----------------------------------------------------------------------------------------------------------------------------------------------
'Action Name : Add_To_Favorites_By_Icon
'Description : This action is to add video to favorites on Workout Videos subpage by Heart icon
'Creation Date : 24.10.2018
'Author: Natalia Welter
'Last modification date : None
'Outputs: strVideoLinkID
'Assumptions /Effects : Workout Video is added to favorites
'Returns : Action returns 0 if everything is correct or returns ActionNumber + step number if error occurs
'-----------------------------------------------------------------------------------------------------------------------------------------------
Option Explicit
Reporter.Filter = rfDisableAll

' Hover over Workout Video from the list
If Browser("WebBrowser").Page("WorkoutVideos").WebElement("WorkoutVideosSection").Link("VideoSecond").Exist(5) Then
	Setting.WebPackage("ReplayType") = 2
	Browser("WebBrowser").Page("WorkoutVideos").WebElement("WorkoutVideosSection").Link("VideoSecond").FireEvent("onmouseover")
	
	'Get ID from video link to <<strVideoLinkID>>
	Parameter ("strVideoLinkID") = Trim(Browser("WebBrowser").Page("WorkoutVideos").WebElement("WorkoutVideosSection").Link("VideoSecond").GetROProperty("html id"))
	
	'Click on Heart icon
	If Browser("WebBrowser").Page("WorkoutVideos").WebElement("WorkoutVideosSection").Link("VideoSecond").WebButton("HeartIcon").Exist(5) Then
		Browser("WebBrowser").Page("WorkoutVideos").WebElement("WorkoutVideosSection").Link("VideoSecond").WebButton("HeartIcon").Click
		Setting.WebPackage("ReplayType") = 1
		Browser("WebBrowser").Page("WorkoutVideos").WebElement("WorkoutVideosSection").Object.scrollIntoView
		fnReportStepEx "Pass", "Hover over Workout Video from the list and click on Heart icon.", "Workout Video is selected. Heart icon is selected.", Browser("WebBrowser"), "true"	
	Else
		Browser("WebBrowser").Page("WorkoutVideos").WebElement("WorkoutVideosSection").Object.scrollIntoView
		fnReportStepEx "Fail", "Hover over Workout Video from the list and click on Heart icon.", "Workout Video is selected. Heart icon is NOT selected.", Browser("WebBrowser"), "true"
		ExitActionIteration "Add_To_Favorites_By_Icon.1.1"		
	End If
Else
	Browser("WebBrowser").Page("WorkoutVideos").WebElement("WorkoutVideosSection").Object.scrollIntoView
	fnReportStepEx "Fail", "Hover over Workout Video from the list and click on Heart icon.", "Workout Video is NOT displayed", Browser("WebBrowser"), "true"
	ExitActionIteration "Add_T0_Favorites_By_Icon.1.2"	
End If

ExitActionIteration "0"
