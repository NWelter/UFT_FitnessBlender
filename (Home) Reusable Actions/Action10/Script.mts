'-----------------------------------------------------------------------------------------------------------------------------------------------
'Action Name : Check_SearchingResults
'Description : This action is to compare if searching keyword <<strKeyword>> is contained in text results on Search Results subpage
'Creation Date : 29.10.2018
'Last modification date : None
'Inputs: strKeyword
'Assumptions /Effects : Searching keyword is displayed in text results on Search Results subpage
'Returns : Action returns 0 if everything is correct or returns ActionNumber + step number if error occurs
'-----------------------------------------------------------------------------------------------------------------------------------------------
Option Explicit
Reporter.Filter = rfDisableAll

Dim strNoResult

strNoResult = "Sorry, there are no results for "

'Check the search results for <<strKeyword>>
If Parameter("strKeyword") <> "" Then
	If Browser("WebBrowser").Page("SearchResults").WebElement("Article").Exist(5) Then
		If Browser("WebBrowser").Page("SearchResults").WebElement("Article").fnContains("outertext", Parameter("strKeyword")) Then
			fnReportStepEx "Pass", "Check the search results for keyword " & "'" & Parameter("strKeyword") & "'",_
			"Searching keyword " & "'" & Parameter("strKeyword") & "'" & " is contained in text results."  , Browser("WebBrowser"), "true"
		Else 
			fnReportStepEx "Fail", "Check the search results for keyword " & "'" & Parameter("strKeyword") & "'" ,_ 
			"Searching keyword " & "'" & Parameter("strKeyword") & "'" & " is NOT contained in text results."  , Browser("WebBrowser"), "true"
			ExitActionIteration "Check_SearchingResults.1"
		End If
	Else 
		fnReportStepEx "Fail", "Check the search results for keyword "  & "'" & Parameter("strKeyword") & "'",_ 
		"Search results are NOT displayed", Browser("WebBrowser"), "true"	
	End If
End If

'Check the search results for empty keyword <<strKeyword>>
If Parameter("strKeyword") = "" Then
	If Browser("WebBrowser").Page("SearchResults").WebElement("NoResultsHeader").Exist(5) Then
		If Browser("WebBrowser").Page("SearchResults").WebElement("NoResultsHeader").fnContains("outertext", strNoResult) Then
			fnReportStepEx "Pass", "Check the search results for empty keyword.",_ 
			"'No results' text is displayed: " & "'" & strNoResult & " '" & Parameter("strKeyword")& "'" & " '", Browser("WebBrowser"), "true"
		End If
	Else
		fnReportStepEx "Fail", "Check the search results for empty keyword",_ 
		"Search results are NOT displayed", Browser("WebBrowser"), "true"
		ExitActionIteration "Check_SearchingResults.2"		
	End If
End If

ExitActionIteration "0"


