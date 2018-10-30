'----------------------------------------------------------------------------------------------------------------------------------------------
'Action Name : Login
'Description : This action is to open <<strBrowser>> on <<strURL>> as <<strUserName>> and <<strPassword>>
'Author  : Natalia Welter
'Creation Date : 05.10.2018
'Last modification date : None
'Assumptions /Effects : User correctly opened open <<strBrowser>> on <<strURL>> as <<strUsername>> and <<strPassword>>
'Inputs : strBrowser, strURL, strUserName, strPassword
'Returns : Action returns 0 if everything is correct or returns ActionNumber + step number if error occurs
'-----------------------------------------------------------------------------------------------------------------------------------------------
Option Explicit
Reporter.Filter = rfDisableAll

Dim strRunActionStatus 

'Prepare environment for testing 
'1) - close all browsers
If Not fnCloseBrowserEx (Browser("WebBrowser")) Then
	fnReportStepDesktopScreen "Fail", "Close all browsers", "Browsers are NOT closed"
	ExitActionIteration "Login.1"
Else
	fnReportStepDesktopScreen "Pass", "Close all browsers", "All browsers are closed"
End If

'2) - Open Browser on URL
fnStartBrowserEx Parameter("strBrowser"), Parameter("strURL")

If Browser("WebBrowser").Page("Home").Exist(30) Then
	fnReportStepEx "Pass", "Open '" & Parameter("strBrowser") & "' on URL '" & Parameter("strURL") & "'", "'Browser correctly opened", Browser("WebBrowser"), "true"
	fnReportStepDesktopScreen "Pass", "Open '" & Parameter("strBrowser") & "' on URL '" & Parameter("strURL") & "'", "'Browser correctly opened"
Else
	fnReportStepEx "Fail", "Open '" & Parameter("strBrowser") & "' on URL '" & Parameter("strURL") & "'", "'Browser NOT correctly opened", Browser("WebBrowser"), "true"
	ExitActionIteration "Login.2"
End If

ExitActionIteration "0"
