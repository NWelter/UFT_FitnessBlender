'------------------------------------------------------------------------------------------------------------
'Action Name: Search
'Description: This action is to set search value <<strKeywords>> and display search results
'Creation Date: 08-10-2018
'Last modification date: <None>
'Assumptions / Effects: Search Results subpage is displayed
'Returns: Action return 0 if everything is correct or returns ActionNumber + step number if error occure
'------------------------------------------------------------------------------------------------------------

'Click on scope icon
If fnFireDownClick(Browser("WebBrowser").Page("AllPages").WebElement("ScopeIcon")) Then
	fnReportStepEx "Pass", "Click on scope icon", "Scope icon is extended", Browser("WebBrowser"), "true"
Else
	fnReportStepEx "Fail", "Click on scope icon", "Scope icon is NOT extended", Browser("WebBrowser"), "true"
	ExitActionIteration "Search.1"
End If

'Set search value <<strKeywords>>
If fnSet(Browser("WebBrowser").Page("AllPages").WebEdit("Keyword"), Parameter("strKeywords")) Then
	fnReportStepEx "Pass", "Set value '" & Parameter("strKeywords") & "'", "Value '" & Parameter("strKeywords") & "' is set correctly", Browser("WebBrowser"), "true"
Else
	fnReportStepEx "Fail", "Set value '" & Parameter("strKeywords") & "'", "Value '" & Parameter("strKeywords") & "' is NOT set correctly", Browser("WebBrowser"), "true"
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

'If fnContains(browser("WebBrowser").Page("SearchResults").WebElement("SearchHeader"), "outertext", Parameter("strKeywords")) OR fnContains(Browser("WebBrowser").Page("SearchResults").WebElement("Article"), "outertext", Parameter("strKeywords")) Then
'	fnReportStepEx "Pass", "Check the search results", "Search results contain search keywords <<strKeywords>>", Browser("WebBrowser"), "true"
'Else
'	fnReportStepEx "Fail", "Check the search results", "Search results contain search keywords <<strKeywords>>", Browser("WebBrowser"), "true"
'End  If 

ExitActionIteration "0"



 @@ script infofile_;_ZIP::ssf13.xml_;_
