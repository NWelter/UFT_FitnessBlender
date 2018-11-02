'------------------------------------------------------------------------------------------------------------
'Action Name: AddingWorkoutVideoToFavorites
'Description: This action is to verify that logged user can add workout video to favorites
'Creation Date: 23-10-2018
'Author: Natalia Welter
'Last modification date: <None>
'Assumptions / Effects: Logged user added workout video to favorites succesfully
'Inputs: strBrowser, strURL, strUsername, strPassword
'Returns: Action return 0 if everything is correct or returns ActionNumber + step number if error occure
'------------------------------------------------------------------------------------------------------------
Option Explicit
Reporter.Filter = rfDisableAll

Dim strRunActionStatus, blnControlFlow

blnControlFlow = True

' Step 1. Open <<browser>> and go to <<URL>>
strRunActionStatus = "9999"
strRunActionStatus = RunAction ("Login [(Main) Reusable Actions]", oneIteration, Parameter("strBrowser"), Parameter("strURL"))
If (StrComp(strRunActionStatus, "0", 1) <> 0) Then
    fnReportStep "Fail", "Login action failed" , "Returned value: " & strRunActionStatus , ""
    ExitActionIteration "AddingWorkoutVideoToFavorites.1"
End If

' Step 2. Hover over My Fitness dropdown. Click on Sign In button.
strRunActionStatus = "9999"
strRunActionStatus = RunAction ("Navigate_SignIn [(Main) Reusable Actions]", oneIteration)
If (StrComp(strRunActionStatus, "0", 1) <> 0) Then
    fnReportStep "Fail", "Navigate_SignIn action failed" , "Returned value: " & strRunActionStatus , ""
    ExitActionIteration "AddingWorkoutVideoToFavorites.2"
End If

' Step 3. Set login and password on login form fields. Click on Sign In button.
strRunActionStatus = "9999"
strRunActionStatus = RunAction ("SignIn [(Main) Reusable Actions]", oneIteration, Parameter("strUsername"), Parameter("strPassword"))
If (StrComp(strRunActionStatus, "0", 1) <> 0) Then
    fnReportStep "Fail", "SignIn action failed" , "Returned value: " & strRunActionStatus , ""
    ExitActionIteration "AddingWorkoutVideoToFavorites.3"
End If

' Step 4. Hover over Workouts&Programs dropdown on header navbar. Click on Workout Videos subtab.
strRunActionStatus = "9999"
strRunActionStatus = RunAction ("Navigate_WorkoutVideos [(WorkoutVideos) Reusable Actions]", oneIteration)
If (StrComp(strRunActionStatus, "0", 1) <> 0) Then
    fnReportStep "Fail", "Navigate_WorkoutVideos action failed" , "Returned value: " & strRunActionStatus , ""
    ExitActionIteration "AddingWorkoutVideoToFavorites.4"
End If

' Step 5. Click on first workout video from list.
strRunActionStatus = "9999"
strRunActionStatus = RunAction ("Select_Video [(WorkoutVideos) Reusable Actions]", oneIteration)
If (StrComp(strRunActionStatus, "0", 1) <> 0) Then
    fnReportStep "Fail", "Select_Video action failed" , "Returned value: " & strRunActionStatus , ""
    blnControlFlow = False
End If

If blnControlFlow Then
	' Step 6. Click on Add To Favorites button.
	strRunActionStatus = "9999"
	strRunActionStatus = RunAction ("Add_To_Favorites [(WorkoutVideosDetails)Reusable Actions]", oneIteration)
	If (StrComp(strRunActionStatus, "0", 1) <> 0) Then
	    fnReportStep "Fail", "Add_To_Favorites action failed" , "Returned value: " & strRunActionStatus , ""
	End If
	
	' Step 7. Hover over My Fitness dropdown on header navbar.
	strRunActionStatus = "9999"
	strRunActionStatus = RunAction ("Hover_MyFitnessLoggedUser [(Home) Reusable Actions]", oneIteration)
	If (StrComp(strRunActionStatus, "0", 1) <> 0) Then
	    fnReportStep "Fail", "Hover_MyFitnessLoggedUser action failed" , "Returned value: " & strRunActionStatus , ""
	End If
	
	' Step 8. Click on Favorites subtab.
	strRunActionStatus = "9999"
	strRunActionStatus = RunAction ("Navigate_Favorites [(Favorites) Reusable Actions]", oneIteration)
	If (StrComp(strRunActionStatus, "0", 1) <> 0) Then
	    fnReportStep "Fail", "Navigate_Favorites action failed" , "Returned value: " & strRunActionStatus , ""
	End If
	
	' Check if proper video link is added to favorites
	strRunActionStatus = "9999"
	strRunActionStatus = RunAction ("Check_VideoID [(Favorites) Reusable Actions]", oneIteration, Parameter("Select_Video [(WorkoutVideos) Reusable Actions]", "strVideoLinkID"))
	If (StrComp(strRunActionStatus, "0", 1) <> 0) Then
	    fnReportStep "Fail", "Check_VideoID action failed" , "Returned value: " & strRunActionStatus , ""
	End If	
End If

' Step 9. Hover over Workouts&Programs dropdown on header navbar. Click on Workout Videos subtab.
strRunActionStatus = "9999"
strRunActionStatus = RunAction ("Navigate_WorkoutVideos [(WorkoutVideos) Reusable Actions]", oneIteration)
If (StrComp(strRunActionStatus, "0", 1) <> 0) Then
    fnReportStep "Fail", "Navigate_WorkoutVideos action failed" , "Returned value: " & strRunActionStatus , ""
    ExitActionIteration "AddingWorkoutVideoToFavorites.5"
End If

' Step 10. Hover over second workout video from list. Click on Heart icon.
strRunActionStatus = "9999"
strRunActionStatus = RunAction ("Add_To_Favorites_By_Icon [(WorkoutVideos) Reusable Actions]", oneIteration, Parameter("strIconColor"), Parameter("strBorderColor"))
If (StrComp(strRunActionStatus, "0", 1) <> 0) Then
    fnReportStep "Fail", "Add_To_Favorites_By_Icon action failed" , "Returned value: " & strRunActionStatus , ""
    ExitActionIteration "AddingWorkoutVideoToFavorites.6"
End If

' Step 11. Hover over My Fitness dropdown on header navbar.
strRunActionStatus = "9999"
strRunActionStatus = RunAction ("Hover_MyFitnessLoggedUser [(Home) Reusable Actions]", oneIteration)
If (StrComp(strRunActionStatus, "0", 1) <> 0) Then
    fnReportStep "Fail", "Hover_MyFitnessLoggedUser action failed" , "Returned value: " & strRunActionStatus , ""
End If

'  Click on Favorites subtab.
strRunActionStatus = "9999"
strRunActionStatus = RunAction ("Navigate_Favorites [(Favorites) Reusable Actions]", oneIteration)
If (StrComp(strRunActionStatus, "0", 1) <> 0) Then
    fnReportStep "Fail", "Navigate_Favorites action failed" , "Returned value: " & strRunActionStatus , ""
End If

' Check if proper video link is added to favorites
strRunActionStatus = "9999"
strRunActionStatus = RunAction ("Check_VideoID [(Favorites) Reusable Actions]", oneIteration, Parameter("Add_To_Favorites_By_Icon [(WorkoutVideos) Reusable Actions]", "strVideoLinkID"))
If (StrComp(strRunActionStatus, "0", 1) <> 0) Then
    fnReportStep "Fail", "Check_VideoID action failed" , "Returned value: " & strRunActionStatus , ""
End If

ExitActionIteration "0"

