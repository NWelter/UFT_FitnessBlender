'------------------------------------------------------------------------------------------------------------
'Action Name: AddingSingleProductToCart
'Description: This action is to verify that logged user can add single product to Cart
'Creation Date: 02-11-2018
'Author: Natalia Welter
'Last modification date: <None>
'Assumptions / Effects: Logged user added single product to Cart succesfully
'Inputs: strBrowser, strURL, strUsername, strPassword
'Returns: Action return 0 if everything is correct or returns ActionNumber + step number if error occure
'------------------------------------------------------------------------------------------------------------
Option Explicit
Reporter.Filter = rfDisableAll

Dim strRunActionStatus

' Step 1. Open <<browser>>  and go to <<URL>>
strRunActionStatus = "9999"
strRunActionStatus = RunAction ("Login [(Main) Reusable Actions]", oneIteration, Parameter("strBrowser"), Parameter("strURL"), Parameter("strUsername"), Parameter("strPassword"))
If (StrComp(strRunActionStatus, "0", 1) <> 0) Then
    fnReportStep "Fail", "Login action failed" , "Returned value: " & strRunActionStatus , ""
    ExitActionIteration "AddingSingleProductToCart.1"
End If

' Step 2. Hover over My Fitness dropdown. Click on Sign In button.
strRunActionStatus = "9999"
strRunActionStatus = RunAction ("Navigate_SignIn [(Main) Reusable Actions]", oneIteration)
If (StrComp(strRunActionStatus, "0", 1) <> 0) Then
    fnReportStep "Fail", "Navigate_SignIn action failed" , "Returned value: " & strRunActionStatus , ""
    ExitActionIteration "AddingSingleProductToCart.2"
End If

' Step 3. Set login <<strUsername>> and password <<strPassword>> on login form fields. Click on Sign In button.
strRunActionStatus = "9999"
strRunActionStatus = RunAction ("SignIn [(Main) Reusable Actions]", oneIteration, Parameter("strUsername"), Parameter("strPassword"))
If (StrComp(strRunActionStatus, "0", 1) <> 0) Then
    fnReportStep "Fail", "SignIn action failed" , "Returned value: " & strRunActionStatus , ""
    ExitActionIteration "AddingSingleProductToCart.3"
End If

' Step 4. Hover over Workouts&Programs dropdown on header navbar. Click on Meal Plans subtab.
strRunActionStatus = "9999"
strRunActionStatus = RunAction ("Navigate_MealPlans [(MealPlans) Reusabe Actions]", oneIteration)
If (StrComp(strRunActionStatus, "0", 1) <> 0) Then
    fnReportStep "Fail", "Navigate_MealPlans action failed" , "Returned value: " & strRunActionStatus , ""
    ExitActionIteration "AddingSingleProductToCart.4"
End If

' Step 5. Click on first Meal Plan header link under Calendar Based Plans.
strRunActionStatus = "9999"
strRunActionStatus = RunAction ("Select_CalendarMealPlanByHeader [(MealPlans) Reusabe Actions]", oneIteration)
If (StrComp(strRunActionStatus, "0", 1) <> 0) Then
    fnReportStep "Fail", "Select_CalendarMealPlanByHeader action failed" , "Returned value: " & strRunActionStatus , ""
    ExitActionIteration "AddingSingleProductToCart.5"
End If

' Step 6. Click on Add To Bag button.
strRunActionStatus = "9999"
strRunActionStatus = RunAction ("Add_To_Cart [(MealPlanDetails) Reusable Actions]", oneIteration)
If (StrComp(strRunActionStatus, "0", 1) <> 0) Then
    fnReportStep "Fail", "Add_To_Cart action failed" , "Returned value: " & strRunActionStatus , ""
    ExitActionIteration "AddingSingleProductToCart.6"
End If

strRunActionStatus = "9999"
strRunActionStatus = RunAction ("Check_Item [(Cart) Reusable Actions]", oneIteration, Parameter("Select_CalendarMealPlanByHeader [(MealPlans) Reusabe Actions]", "strMealPlanLink"))
If (StrComp(strRunActionStatus, "0", 1) <> 0) Then
    fnReportStep "Fail", "Check_Item action failed" , "Returned value: " & strRunActionStatus , ""
    ExitActionIteration "AddingSingleProductToCart.7"
End If

strRunActionStatus = "9999"
strRunActionStatus = RunAction ("Compare_ItemsAmount [(Cart) Reusable Actions]", oneIteration)
If (StrComp(strRunActionStatus, "0", 1) <> 0) Then
    fnReportStep "Fail", "Compare_ItemsAmount action failed" , "Returned value: " & strRunActionStatus , ""
    ExitActionIteration "AddingSingleProductToCart.8"
End If

ExitActionIteration "0"



