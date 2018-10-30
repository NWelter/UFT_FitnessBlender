'------------------------------------------------------------------------------------------------------------
'Action Name: CreateNewAccount
'Description: This action is to verify that new user is able to registry and sign in
'Creation Date: 25-10-2018
'Author: Natalia Welter
'Last modification date: <None>
'Assumptions / Effects: New user is registered and logged successfully
'Inputs: strFirstName, strLastName, strUsername, strPassword
'Returns: Action return 0 if everything is correct or returns ActionNumber + step number if error occure
'------------------------------------------------------------------------------------------------------------
Option Explicit
Reporter.Filter = rfDisableAll

Dim strRunActionStatus, strInfoHeader, strInvalidHeader

' Step 1. Open <<browser>>  and go to <<URL>>
strRunActionStatus = "9999"
strRunActionStatus = RunAction ("Login [(Main) Reusable Actions]", oneIteration, Parameter("strBrowser"), Parameter("strURL"), Parameter("strUsername"), Parameter("strPassword"))
If (StrComp(strRunActionStatus, "0", 1) <> 0) Then
    fnReportStep "Fail", "Login action failed" , "Returned value: " & strRunActionStatus , ""
    ExitActionIteration "CreateNewAccount.1"
End If

' Step 2. Hover over My Fitness dropdown. Click on Join button.
strRunActionStatus = "9999"
strRunActionStatus = RunAction ("Navigate_Join [(Main) Reusable Actions]", oneIteration)
If (StrComp(strRunActionStatus, "0", 1) <> 0) Then
    fnReportStep "Fail", "Login action failed" , "Returned value: " & strRunActionStatus , ""
    ExitActionIteration "CreateNewAccount.2"
End If

' Step 3. Fill all required fileds. Click on Join button.
strRunActionStatus = "9999"
strRunActionStatus = RunAction ("FillRegisterForm_Join [(Main) Reusable Actions]", oneIteration, Parameter("strFirstName"),_ 
Parameter("strLastName"), Parameter("strUsername"), Parameter("strPassword"), Parameter("strConfirmPassword"))
If (StrComp(strRunActionStatus, "0", 1) <> 0) Then
    fnReportStep "Fail", "Login action failed" , "Returned value: " & strRunActionStatus , ""
    ExitActionIteration "CreateNewAccount.3"
End If

' Verify New Account subpage content
strRunActionStatus = "9999"
strRunActionStatus = RunAction ("Verify_NewAccount [(New Account) Reusable Actions]", oneIteration)
If (StrComp(strRunActionStatus, "0", 1) <> 0) Then
    fnReportStep "Fail", "Verify_NewAccount action failed" , "Returned value: " & strRunActionStatus , ""
    ExitActionIteration "CreateNewAccount.4"
End If

' Step 4. Click on Create Account button.
strRunActionStatus = "9999"
strRunActionStatus = RunAction ("Confirm_NewAccount [(New Account) Reusable Actions]", oneIteration)
If (StrComp(strRunActionStatus, "0", 1) <> 0) Then
    fnReportStep "Fail", "Confirm_NewAccount action failed" , "Returned value: " & strRunActionStatus , ""
    ExitActionIteration "CreateNewAccount.5"
End If

'Check text from information header on User Dashboard subpage
strRunActionStatus = "9999"
strRunActionStatus = RunAction ("Check_UserDashboardHeader [(Main) Reusable Actions]", oneIteration, Parameter("strDashboardHeader"))
If (StrComp(strRunActionStatus, "0", 1) <> 0) Then
    fnReportStep "Fail", "SignOut action failed" , "Returned value: " & strRunActionStatus , ""
End If

' Step 5. Hover over My Fitness dropdown. Click on Sign Out subtab.
strRunActionStatus = "9999"
strRunActionStatus = RunAction ("SignOut [(Main) Reusable Actions]", oneIteration)
If (StrComp(strRunActionStatus, "0", 1) <> 0) Then
    fnReportStep "Fail", "SignOut action failed" , "Returned value: " & strRunActionStatus , ""
    ExitActionIteration "CreateNewAccount.6"
End If

' Step 6. Hover over My Fitness dropdown. Click on Sign In button.
strRunActionStatus = "9999"
strRunActionStatus = RunAction ("Navigate_SignIn [(Main) Reusable Actions]", oneIteration)
If (StrComp(strRunActionStatus, "0", 1) <> 0) Then
    fnReportStep "Fail", "Navigate_SignIn action failed" , "Returned value: " & strRunActionStatus , ""
    ExitActionIteration "CreateNewAccount.7"
End If

' Step 7. Type Username and Password which was used during registering a New User. Click on Sign In button.
strRunActionStatus = "9999"
strRunActionStatus = RunAction ("SignIn [(Main) Reusable Actions]", oneIteration, Parameter("FillRegisterForm_Join [(Main) Reusable Actions]", "strUsernameRandomize"), Parameter("strPassword"))
If (StrComp(strRunActionStatus, "0", 1) <> 0) Then
    fnReportStep "Fail", "SignIn action failed" , "Returned value: " & strRunActionStatus , ""
    ExitActionIteration "CreateNewAccount.8"
End If

ExitActionIteration "0"


