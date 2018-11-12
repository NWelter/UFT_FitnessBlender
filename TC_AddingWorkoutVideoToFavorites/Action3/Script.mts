'------------------------------------------------------------------------------------------------------------
'Action Name: SetUp
'Description: This action is to prepare initial state before test case
'Creation Date: 06-11-2018
'Author: Natalia Welter
'Last modification date: <None>
'Assumptions / Effects: Logged user doesn't have added videos on Favorites subpage
'Returns: Action return 0 if everything is correct or returns ActionNumber + step number if error occure
'------------------------------------------------------------------------------------------------------------
Option Explicit
Reporter.Filter = rfDisableAll

Dim strRunActionStatus, objVideoLinks, objObject, strHeaderInfo, i

' 1.) Open <<browser>> and go to <<URL>>
strRunActionStatus = "9999"
strRunActionStatus = RunAction ("Login [(Main) Reusable Actions]", oneIteration, Parameter("strBrowser"), Parameter("strURL"), Parameter("strUsername"), Parameter("strPassword"))
If (StrComp(strRunActionStatus, "0", 1) <> 0) Then
    fnReportStep "Fail", "Login action failed" , "Returned value: " & strRunActionStatus , ""
End If

' 2.) Hover over My Fitness dropdown. Click on Sign In button.
strRunActionStatus = "9999"
strRunActionStatus = RunAction ("Navigate_SignIn [(Main) Reusable Actions]", oneIteration)
If (StrComp(strRunActionStatus, "0", 1) <> 0) Then
    fnReportStep "Fail", "Navigate_SignIn action failed" , "Returned value: " & strRunActionStatus , ""
End If

' 3.) Set login <<strUsername>> and password <<strPassword>> on login form fields. Click on Sign In button.
strRunActionStatus = "9999"
strRunActionStatus = RunAction ("SignIn [(Main) Reusable Actions]", oneIteration, Parameter("strUsername"), Parameter("strPassword"))
If (StrComp(strRunActionStatus, "0", 1) <> 0) Then
    fnReportStep "Fail", "SignIn action failed" , "Returned value: " & strRunActionStatus , ""
End If

' 4.) Hover over My Fitness dropdown on header navbar.
strRunActionStatus = "9999"
strRunActionStatus = RunAction ("Hover_MyFitnessLoggedUser [(Home) Reusable Actions]", oneIteration)
If (StrComp(strRunActionStatus, "0", 1) <> 0) Then
    fnReportStep "Fail", "Hover_MyFitnessLoggedUser action failed" , "Returned value: " & strRunActionStatus , ""
End If

' 5.) Click on Favorites subtab.
strRunActionStatus = "9999"
strRunActionStatus = RunAction ("Navigate_Favorites [(Favorites) Reusable Actions]", oneIteration)
If (StrComp(strRunActionStatus, "0", 1) <> 0) Then
    fnReportStep "Fail", "Navigate_Favorites action failed" , "Returned value: " & strRunActionStatus , ""
End If

' 6.) Remove all added videos from Favorites subpage
	Set objObject = Description.Create()
		objObject("micclass").value = "Link"	
	Set objVideoLinks = Browser("WebBrowser").Page("Favorites").WebElement("FavoriteVideosSection").ChildObjects(objObject)
	
	For i = 0 To objVideoLinks.Count -1
		Browser("WebBrowser").Page("Favorites").WebElement("FavoriteVideosSection").Link("FavoriteVideo").WebButton("HeartIcon").Click
		Browser("WebBrowser").Refresh
	Next
	
	strHeaderInfo = Trim(Browser("WebBrowser").Page("Favorites").WebElement("FavoriteVideosSection").WebElement("VideoHeader").GetROProperty("outertext"))
	
	If strHeaderInfo = "No videos were found" Then
		fnReportStepEx"Pass", "Remove all added videos from Favorites subpage", "All videos are removed", Browser("WebBrowser"), "true"
	Else
		fnReportStepEx "Fail", "Remove all added videos from Favorites subpage", "All videos are NOT removed", Browser("WebBrowser"), "true"
	End If

ExitActionIteration "0"

