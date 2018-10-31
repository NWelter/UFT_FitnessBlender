'-----------------------------------------------------------------------------------------------------------------------------------------------
'Action Name : Add_To_Favorites
'Description : This action is to add video to favorites on Workout Videos details subpage
'Creation Date : 24.10.2018
'Author: Natalia Welter
'Last modification date : None
'Outputs: strVideoDetailsHeader
'Assumptions /Effects : Workout Video is added to favorites
'Returns : Action returns 0 if everything is correct or returns ActionNumber + step number if error occurs
'-----------------------------------------------------------------------------------------------------------------------------------------------
Option Explicit
Reporter.Filter = rfDisableAll

Dim strRunActionStatus, strAdded

'Verify Workout Video details subpage content
strRunActionStatus = "9999"
strRunActionStatus = RunAction ("Verify_WorkoutVideosDetails", oneIteration)
If (StrComp(strRunActionStatus, "0", 1) <> 0) Then
    fnReportStep "Fail", "Verify_WorkoutVideosDetails action failed" , "Returned value: " & strRunActionStatus , ""
End If

'Click on Add To Favorites button
If Browser("WebBrowser").Page("WorkoutVideosDetails").WebButton("AddToFavorites").Exist(5) Then
	Browser("WebBrowser").Page("WorkoutVideosDetails").WebButton("AddToFavorites").Click
	strAdded = Trim(Browser("WebBrowser").Page("WorkoutVideosDetails").WebButton("AddToFavorites").GetROProperty("outertext"))
Else
	fnReportStepEx "Fail", "Click on Add To Favorites button", "Button is NOT displayed", Browser("WebBrowser"), "true"
	ExitActionIteration "Add_To_Favorites.1"
End If

If strAdded = "ADDED TO FAVORITES" Then
	fnReportStepEx "Pass", "Click on Add To Favorites button", "Button is changed correctly."_ 
	& VbCrLf & "Current button title is: " & strAdded , Browser("WebBrowser"), "true"
Else 
	fnReportStepEx "Fail", "Click on Add To Favorites button", "Button is NOT changed correctly."_ 
	& VbCrLf & "Current button title is: " & strAdded, Browser("WebBrowser"), "true"
	ExitActionIteration "Add_To_Favorites.2"
End If

'Get text from video header to <<strVideoDetailsHeader>>
Parameter ("strVideoDetailsHeader") = Trim(Browser("WebBrowser").Page("WorkoutVideosDetails").WebElement("VideoHeader").GetROProperty("outertext"))

ExitActionIteration "0"
