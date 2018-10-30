'------------------------------------------------------------------------------------------------------------
'Action Name: Search
'Description: This action is to set search value <<strKeyword>> and display search results
'Creation Date: 08-10-2018
'Last modification date: <None>
'Assumptions / Effects: Search Results subpage is displayed
'Inputs: strKeyword
'Returns: Action return 0 if everything is correct or returns ActionNumber + step number if error occure
'------------------------------------------------------------------------------------------------------------

Option Explicit
Reporter.Filter = rfDisableAll

'Click on scope icon
If Browser("WebBrowser").Page("AllPages").WebElement("ScopeIcon").fnFireEvent("onmousedown") Then
	fnReportStepEx "Pass", "Click on scope icon", "Scope icon is extended", Browser("WebBrowser"), "true"
Else
	fnReportStepEx "Fail", "Click on scope icon", "Scope icon is NOT extended", Browser("WebBrowser"), "true"
	ExitActionIteration "Search.1"
End If

'Set search value <<strKeyword>>
If fnSet(Browser("WebBrowser").Page("AllPages").WebEdit("Keyword"), Parameter("strKeyword")) Then
	fnReportStepEx "Pass", "Set value '" & Parameter("strKeyword") & "'", "Value '" & Parameter("strKeyword") & "' is set correctly", Browser("WebBrowser"), "true"
Else
	fnReportStepEx "Fail", "Set value '" & Parameter("strKeyword") & "'", "Value '" & Parameter("strKeyword") & "' is NOT set correctly", Browser("WebBrowser"), "true"
	ExitActionIteration "Search.2"
End If

'Click on scope button
If Browser("WebBrowser").Page("AllPages").WebElement("FitnessBlenderHeader").WebButton("Search").Exist(5) Then
	Browser("WebBrowser").Page("AllPages").WebElement("FitnessBlenderHeader").WebButton("Search").Click
	fnReportStepEx "Pass", "Click on scope button", "Scope button is displayed", Browser("WebBrowser"), "true"
Else
	fnReportStepEx "Fail", "Click on scope button", "Scope button is NOT displayed", Browser("WebBrowser"), "true"
	ExitActionIteration "Search.3"
End If


RunAction "Check_SearchingResults", oneIteration, ""

ExitActionIteration "0"



 @@ script infofile_;_ZIP::ssf13.xml_;_
