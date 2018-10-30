'-----------------------------------------------------------------------------------------------------------------------------------------------
'Action Name : Verify_About
'Description : This action is to verify that About subpage is displayed and verify content
'Creation Date : 11.10.2018
'Last modification date : None
'Assumptions /Effects : About subpage and content are correctly displayed
'Returns : Action returns 0 if everything is correct or returns ActionNumber + step number if error occurs
'-----------------------------------------------------------------------------------------------------------------------------------------------
Option Explicit
Reporter.Filter = rfDisableAll

Dim arrPageElementsTop, arrPageElementsMiddle, arrPageElementsBottom, arrCheckResults

' Verify About subpage top sections content
arrPageElementsTop = Array(Browser("WebBrowser").Page("About").WebElement("AboutImage"),_ 
						Browser("WebBrowser").Page("About").WebElement("AboutText"))

arrCheckResults = fnCheckPageElements(arrPageElementsTop)

If arrCheckResults(0) Then
	Browser("WebBrowser").Page("About").WebElement("AboutText").Object.scrollIntoView
	fnReportStepEx "Pass", "Click on About tab." & VbCrLf & "Verify top section content.", "About subpage is displayed." & VbCrLf & "Current elements on top are displayed: " & arrCheckResults(1), Browser("WebBrowser"), "true"
Else 
	fnReportStepEx "Fail", "Click on About tab. Verify top section content.", "About subpage is NOT displayed." & VbCrLf & " Current elements on top are NOT displayed: " & arrCheckResults(2), Browser("WebBrowser"), "true"
	ExitActionIteration "Verify_About.1"
End If

'Verify About subpage middle sections content
arrPageElementsMiddle = Array(Browser("WebBrowser").Page("About").WebElement("VideoImage"),_ 
							Browser("WebBrowser").Page("About").WebElement("FeatureText"))

arrCheckResults = fnCheckPageElements(arrPageElementsMiddle)

If arrCheckResults(0) Then
	Browser("WebBrowser").Page("About").WebElement("FeatureText").Object.scrollIntoView
	fnReportStepEx "Pass", "Click on About tab." & VbCrLf & "Verify middle section content.", "About subpage is displayed." & VbCrLf & "Current elements in the middle are displayed: " & arrCheckResults(1), Browser("WebBrowser"), "true"
Else 
	fnReportStepEx "Fail", "Click on About tab." & VbCrLf & "Verify middle section content.", "About subpage is NOT displayed." & VbCrLf & " Current elements in the middle are NOT displayed: " & arrCheckResults(2), Browser("WebBrowser"), "true"
	ExitActionIteration "Verify_About.2"
End If

' Verify About subpage bottom section content
arrPageElementsBottom = Array(Browser("WebBrowser").Page("About").WebElement("FeatureImageGrid"),_
							Browser("WebBrowser").Page("About").WebElement("InfoText"),_
							Browser("WebBrowser").Page("About").WebElement("BioImage"),_ 
							Browser("WebBrowser").Page("About").WebElement("Mentions"))

arrCheckResults = fnCheckPageElements(arrPageElementsBottom)

If arrCheckResults(0) Then
	Browser("WebBrowser").Page("About").WebElement("InfoText").Object.scrollIntoView
	fnReportStepEx "Pass", "Click on About tab." & VbCrLf &  "Verify bottom section content.", "About subpage is displayed." & VbCrLf & "Current elements on the bottom are displayed: " & arrCheckResults(1), Browser("WebBrowser"), "true"
Else 
	Browser("WebBrowser").Page("About").WebElement("InfoText").Object.scrollIntoView
	fnReportStepEx "Fail", "Click on About tab." & VbCrLf &  "Verify bottom section content.", "About subpage is NOT displayed." & VbCrLf & " Current elements on the bottom are NOT displayed: " & arrCheckResults(2), Browser("WebBrowser"), "true"
	ExitActionIteration "Verify_About.3"
End If

ExitActionIteration "0"

