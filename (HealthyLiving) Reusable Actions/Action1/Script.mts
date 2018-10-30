'-----------------------------------------------------------------------------------------------------------------------------------------------
'Action Name : Verify_HealthyLiving
'Description : This action is to verify that Healthy Living subpage is displayed and verify content
'Creation Date : 10.10.2018
'Last modification date : None
'Assumptions /Effects : Healthy Living subpage and content are correctly displayed
'Returns : Action returns 0 if everything is correct or returns ActionNumber + step number if error occurs
'-----------------------------------------------------------------------------------------------------------------------------------------------
Option Explicit
Reporter.Filter = rfDisableAll

Dim arrPageElementsTop, arrPageElementsMiddle, arrPageElementsBottom, arrCheckResults

' Verify Healthy Living subpage top section content
arrPageElementsTop = Array(Browser("WebBrowser").Page("HealthyLiving").WebElement("FitnessHeader"),_ 
						Browser("WebBrowser").Page("HealthyLiving").WebElement("FitnessSection"))
						
arrCheckResults = fnCheckPageElements(arrPageElementsTop)

If arrCheckResults(0) Then
	Browser("WebBrowser").Page("HealthyLiving").WebElement("FitnessSection").Object.scrollIntoView
	fnReportStepEx "Pass", "Click on Healthy Living dropdown. Verify top section content.", "Healthy Living subpage is displayed." & VbCrLf & "Current sections are available: " & arrCheckResults(1), Browser("WebBrowser"), "true"
Else 
	Browser("WebBrowser").Page("HealthyLiving").WebElement("FitnessSection").Object.scrollIntoView
	fnReportStepEx "Fail", "Click on Healthy Living dropdown. Verify top section content.", "Healthy Living subpage is NOT displayed." & VbCrLf & "Current sections are NOT available: " & arrCheckResults(2), Browser("WebBrowser"), "true"
	ExitActionIteration "Verify_HealthyLiving.1"
End If

' Verify Healthy Living subpage middle section content
arrPageElementsMiddle = Array(Browser("WebBrowser").Page("HealthyLiving").WebElement("HealthHeader"),_
							Browser("WebBrowser").Page("HealthyLiving").WebElement("HealthSection"),_ 
							Browser("WebBrowser").Page("HealthyLiving").WebElement("HealthyRecipesHeader"),_ 
							Browser("WebBrowser").Page("HealthyLiving").WebElement("HealthyRecipesSection"))
						
arrCheckResults = fnCheckPageElements(arrPageElementsMiddle)

If arrCheckResults(0) Then
	Browser("WebBrowser").Page("HealthyLiving").WebElement("HealthSection").Object.scrollIntoView
	fnReportStepEx "Pass", "Click on Healthy Living dropdown. Verify middle section content.", "Healthy Living subpage is displayed." & VbCrLf & "Current sections are available: " & arrCheckResults(1), Browser("WebBrowser"), "true"
Else 
	Browser("WebBrowser").Page("HealthyLiving").WebElement("HealthSection").Object.scrollIntoView
	fnReportStepEx "Fail", "Click on Healthy Living dropdown. Verify middle section content.", "Healthy Living subpage is NOT displayed." & VbCrLf & "Current sections are NOT available: " & arrCheckResults(2), Browser("WebBrowser"), "true"
	ExitActionIteration "Verify_HealthyLiving.2"
End If

' Verify Healthy Living subpage bottom section content
arrPageElementsBottom = Array(Browser("WebBrowser").Page("HealthyLiving").WebElement("Before&AfterHeader"),_ 
						Browser("WebBrowser").Page("HealthyLiving").WebElement("Before&AfterSection"),_
						Browser("WebBrowser").Page("HealthyLiving").WebElement("WeightLossHeader"),_
						Browser("WebBrowser").Page("HealthyLiving").WebElement("WeightLossSection"))
						
arrCheckResults = fnCheckPageElements(arrPageElementsBottom)

If arrCheckResults(0) Then
	Browser("WebBrowser").Page("HealthyLiving").WebElement("Before&AfterSection").Object.scrollIntoView
	fnReportStepEx "Pass", "Click on Healthy Living dropdown. Verify bottom section content.", "Healthy Living subpage is displayed." & VbCrLf & "Current sections are available: " & arrCheckResults(1), Browser("WebBrowser"), "true"
Else 
	Browser("WebBrowser").Page("HealthyLiving").WebElement("Before&AfterSection").Object.scrollIntoView
	fnReportStepEx "Fail", "Click on Healthy Living dropdown. Verify bottom section content.", "Healthy Living subpage is NOT displayed." & VbCrLf & "Current sections are NOT available: " & arrCheckResults(2), Browser("WebBrowser"), "true"
	ExitActionIteration "Verify_HealthyLiving.3"
End If

ExitActionIteration "0"
