'------------------------------------------------------------------------------------------------------------
'Action Name: NavigateToSubpages
'Description: This action is to verify that all navigation tabs content displyed correctly
'Creation Date: 12-10-2018
'Author: Natalia Welter
'Last modification date: <None>
'Assumptions / Effects: User can navigate to all subpages available in navigation tabs. Every link is displaying appropriate subpage.
'Returns: Action return 0 if everything is correct or returns ActionNumber + step number if error occure
'------------------------------------------------------------------------------------------------------------
Option Explicit
Reporter.Filter = rfDisableAll

Dim strRunActionStatus, strCartHeader, strCurrentHeaderText

' Step 1: Open <<browser>> and go to <<URL>>
strRunActionStatus = "9999"
strRunActionStatus = RunAction ("Login [(Main) Reusable Actions]", oneIteration, Parameter("strBrowser"), Parameter("strURL"), Parameter("strUsername"), Parameter("strPassword"))
If (StrComp(strRunActionStatus, "0", 1) <> 0) Then
    fnReportStep "Fail", "Login action failed" , "Returned value: " & strRunActionStatus , ""
    ExitActionIteration "NavigateToSubpages.1"
End If

strRunActionStatus = "9999"
strRunActionStatus = RunAction ("Verify_FitnessBlenderHeader [(Home) Reusable Actions]", oneIteration)
If (StrComp(strRunActionStatus, "0", 1) <> 0) Then
    fnReportStep "Fail", "Verify_FitnessBlenderHeader action failed" , "Returned value: " & strRunActionStatus , ""
    ExitActionIteration "NavigateToSubpages.2"
End If

' Step 2: Click on Workouts&Programs dropdown.
strRunActionStatus = "9999"
strRunActionStatus = RunAction ("Navigate_WorkoutsAndPrograms [(Workouts&Programs) Reusable Actions]", oneIteration)
If (StrComp(strRunActionStatus, "0", 1) <> 0) Then
    fnReportStep "Fail", "Navigate_WorkoutsAndPrograms action failed" , "Returned value: " & strRunActionStatus , ""
End If

' Step 3: Hover over Workouts & Programs dropdown.
strRunActionStatus = "9999"
strRunActionStatus = RunAction ("Hover_WorkoutsAndProgramsDropdown [(Home) Reusable Actions]", oneIteration)
If (StrComp(strRunActionStatus, "0", 1) <> 0) Then
    fnReportStep "Fail", "Hover_WorkoutsAndProgramsDropdown action failed" , "Returned value: " & strRunActionStatus , ""
End If

' Step 4: Click on Workout Videos subtab.
strRunActionStatus = "9999"
strRunActionStatus = RunAction ("Navigate_WorkoutVideos [(WorkoutVideos) Reusable Actions]", oneIteration)
If (StrComp(strRunActionStatus, "0", 1) <> 0) Then
    fnReportStep "Fail", "Navigate_WorkoutVideos action failed" , "Returned value: " & strRunActionStatus , ""
End If

' Step 5: Hover over Workouts & Programs dropdown. Click on Workout Programs subtab.
strRunActionStatus = "9999"
strRunActionStatus = RunAction ("Navigate_WorkoutPrograms [(WorkoutPrograms) Reusable Actions]", oneIteration)
If (StrComp(strRunActionStatus, "0", 1) <> 0) Then
    fnReportStep "Fail", "Navigate_WorkoutPrograms action failed" , "Returned value: " & strRunActionStatus , ""
End If

' Step 6: Hover over Workouts & Programs dropdown. Click on Meal Plans subtab.
strRunActionStatus = "9999"
strRunActionStatus = RunAction ("Navigate_MealPlans [(MealPlans) Reusabe Actions]", oneIteration)
If (StrComp(strRunActionStatus, "0", 1) <> 0) Then
    fnReportStep "Fail", "Navigate_MealPlans action failed" , "Returned value: " & strRunActionStatus , ""
End If

' Step 7: Click on Healthy Living dropdown.
strRunActionStatus = "9999"
strRunActionStatus = RunAction ("Navigate_HealthyLiving [(HealthyLiving) Reusable Actions]", oneIteration)
If (StrComp(strRunActionStatus, "0", 1) <> 0) Then
    fnReportStep "Fail", "Navigate_HealthyLiving action failed" , "Returned value: " & strRunActionStatus , ""
End If

' Step 8: Hover over Healthy Living dropdown
strRunActionStatus = "9999"
strRunActionStatus = RunAction ("Hover_HealthyLivingDropdown [(Home) Reusable Actions]", oneIteration)
If (StrComp(strRunActionStatus, "0", 1) <> 0) Then
    fnReportStep "Fail", "Hover_HealthyLivingDropdown action failed" , "Returned value: " & strRunActionStatus , ""
End If

' Step 9: Click on Fitness subtab.
strRunActionStatus = "9999"
strRunActionStatus = RunAction ("Navigate_Fitness [(Fitness) Reusable Actions]", oneIteration)
If (StrComp(strRunActionStatus, "0", 1) <> 0) Then
    fnReportStep "Fail", "Navigate_Fitness action failed" , "Returned value: " & strRunActionStatus , ""
End If

' Step 10: Hover over Healthy Living dropdown. Click on Health subtab.
strRunActionStatus = "9999"
strRunActionStatus = RunAction ("Navigate_Health [(Health) Reusable Actions]", oneIteration)
If (StrComp(strRunActionStatus, "0", 1) <> 0) Then
    fnReportStep "Fail", "Navigate_Health action failed" , "Returned value: " & strRunActionStatus , ""
End If

' Step 11: Hover over Healthy Living dropdown. Click on Healthy Recipes subtab.
strRunActionStatus = "9999"
strRunActionStatus = RunAction ("Navigate_HealthyRecipes [(Healthy Recipes) Reusable Actions]", oneIteration)
If (StrComp(strRunActionStatus, "0", 1) <> 0) Then
    fnReportStep "Fail", "Navigate_HealthyRecipes action failed" , "Returned value: " & strRunActionStatus , ""
End If

' Step 12: Hover over Healthy Living dropdown. Click on Before&After subtab.
strRunActionStatus = "9999"
strRunActionStatus = RunAction ("Navigate_BeforeAndAfter [(Before&After) Reusable Actions]", oneIteration)
If (StrComp(strRunActionStatus, "0", 1) <> 0) Then
    fnReportStep "Fail", "Verify_Before&After action failed" , "Returned value: " & strRunActionStatus , ""
End If

' Step 13: Hover over Healthy Living dropdown. Click on Weight Loss subtab.
strRunActionStatus = "9999"
strRunActionStatus = RunAction ("Navigate_WeightLoss [(WeightLoss) Reusable Actions]", oneIteration)
If (StrComp(strRunActionStatus, "0", 1) <> 0) Then
    fnReportStep "Fail", "Navigate_WeightLoss action failed" , "Returned value: " & strRunActionStatus , ""
End If

' Step 14: Click on Community tab.
strRunActionStatus = "9999"
strRunActionStatus = RunAction ("Navigate_Community [(Community) Reusable Actions]", oneIteration)
If (StrComp(strRunActionStatus, "0", 1) <> 0) Then
    fnReportStep "Fail", "Navigate_Community action failed" , "Returned value: " & strRunActionStatus , ""
End If

' Step 15: Click on Blog tab
strRunActionStatus = "9999"
strRunActionStatus = RunAction ("Navigate_Blog [(Blog) Reusable Actions]", oneIteration)
If (StrComp(strRunActionStatus, "0", 1) <> 0) Then
    fnReportStep "Fail", "Navigate_Blog action failed" , "Returned value: " & strRunActionStatus , ""
End If

' Step 16: Click on About tab.
strRunActionStatus = "9999"
strRunActionStatus = RunAction ("Navigate_About [(About) Reusable Actions]", oneIteration)
If (StrComp(strRunActionStatus, "0", 1) <> 0) Then
    fnReportStep "Fail", "Navigate_About action failed" , "Returned value: " & strRunActionStatus , ""
End If

' Step 17: Hover over My Fitness dropdown.
strRunActionStatus = "9999"
strRunActionStatus = RunAction ("Hover_MyFitness [(Home) Reusable Actions]", oneIteration)
If (StrComp(strRunActionStatus, "0", 1) <> 0) Then
    fnReportStep "Fail", "Hover_MyFitness action failed" , "Returned value: " & strRunActionStatus , ""
End If

' Step 18: Click on Join button.
strRunActionStatus = "9999"
strRunActionStatus = RunAction ("Navigate_Join [(Main) Reusable Actions]", oneIteration)
If (StrComp(strRunActionStatus, "0", 1) <> 0) Then
    fnReportStep "Fail", "Navigate_Join action failed" , "Returned value: " & strRunActionStatus , ""
End If

' Step 19: Hover over My Fitness dropdown. Click on Sign In button.
strRunActionStatus = "9999"
strRunActionStatus = RunAction ("Navigate_SignIn [(Main) Reusable Actions]", oneIteration)
If (StrComp(strRunActionStatus, "0", 1) <> 0) Then
    fnReportStep "Fail", "Navigate_SignIn action failed" , "Returned value: " & strRunActionStatus , ""
End If

' Step 20: Click on a Shopping Bag icon at right top corner.
strRunActionStatus = "9999"
strRunActionStatus = RunAction ("Navigate_Cart [(Cart) Reusable Actions]", oneIteration)
If (StrComp(strRunActionStatus, "0", 1) <> 0) Then
    fnReportStep "Fail", "Navigate_Cart action failed" , "Returned value: " & strRunActionStatus , ""
End If

' Check text from the Cart subpge header
strCartHeader = "Shopping Bag is Empty"
strCurrentHeaderText = Trim(Browser("WebBrowser").Page("Cart").WebElement("ShoppingCartHeader").GetROProperty("outertext"))

If strCurrentHeaderText = strCartHeader Then
	fnReportStepEx "Pass", "Check text from the Cart subpage header", "Text from the Cart subpage header is equal to: " & strCartHeader, Browser("WebBrowser"), "true"
Else
	fnReportStepEx "Fail", "Check text from the Cart subpage header", "Text from the Cart subpage header is NOT equal to: " & strCartHeader, Browser("WebBrowser"), "true"
End  If

' Step: 21 Click on Fitness Blender Logo button.
strRunActionStatus = "9999"
strRunActionStatus = RunAction ("Redirect_Home [(Home) Reusable Actions]", oneIteration)
If (StrComp(strRunActionStatus, "0", 1) <> 0) Then
    fnReportStep "Fail", "Redirect_Home action failed" , "Returned value: " & strRunActionStatus , ""
End If

ExitActionIteration "0"



