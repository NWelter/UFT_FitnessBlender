'-----------------------------------------------------------------------------------------------------------------------------------------------
'Action Name : Select_CalendarMealPlanByHeader
'Description : This action is to select Calendar Based Meal Plan link on Meal Plans subpage
'Creation Date : 02.11.2018
'Author: Natalia Welter
'Last modification date : None
'Outputs: strMealPlanHeader
'Assumptions /Effects : Meal Plan details subpage is displayed
'Returns : Action returns 0 if everything is correct or returns ActionNumber + step number if error occurs
'-----------------------------------------------------------------------------------------------------------------------------------------------
Option Explicit
Reporter.Filter = rfDisableAll

Dim strMealPlanPageURL

' Click on Meal Plan header link in Calendar Based Meal Plan section
If Browser("WebBrowser").Page("MealPlans").WebElement("CalendarPlansSection").Link("CalendarMealPlanLink").Exist(5) Then
Parameter ("strMealPlanLink") = Trim(Browser("WebBrowser").Page("MealPlans").WebElement("CalendarPlansSection").Link("CalendarMealPlanLink").GetROProperty("href"))
	Browser("WebBrowser").Page("MealPlans").WebElement("CalendarPlansSection").Link("CalendarMealPlanLink").Click
Else
	fnReportStepEx "Fail", "Click on Meal Plan header link in Calendar Based Meal Plan section", "Meal Plan header link is NOT displayed", Browser("WebBrowser"), "true"
	ExitActionIteration "Select_CalendarMealPlanByHeader.1"
End If

' Check if valid Meal Plan details subpage is opened
strMealPlanPageURL = Trim(Browser("WebBrowser").Page("MealPlanDetails").GetROProperty("url"))
If Parameter ("strMealPlanLink") <> strMealPlanPageURL Then
	fnReportStepEx "Fail", "Check if valid Meal Plan details subpage is opened",_  
	"Valid Meal Plan details subpage is NOT opened." & VbCrLf & "Valid URL is: " & Parameter ("strMealPlanLink") & VbCrLf & "Current page URL is: " & strMealPlanPageURL, Browser("WebBrowser"), "true"
	ExitActionIteration "Select_CalendarMealPlanByHeader.2"	
End If

If Browser("WebBrowser").Page("MealPlanDetails").WebElement("ProgramDetails").Exist(10) Then
	fnReportStepEx "Pass", "Click on Meal Plan header link in Calendar Based Meal Plan section", "Meal Plan details subpage is displayed", Browser("WebBrowser"), "true"
Else 
	fnReportStepEx "Fail", "Click on Meal Plan header link in Calendar Based Meal Plan section", "Meal Plan details subpage is NOT displayed", Browser("WebBrowser"), "true"
	ExitActionIteration "Select_CalendarMealPlanByHeader.3"
End If

' Verify Meal Plan details subpage content
RunAction "Verify_MealPlanDetails [(MealPlanDetails) Reusable Actions]", oneIteration

ExitActionIteration "0"


