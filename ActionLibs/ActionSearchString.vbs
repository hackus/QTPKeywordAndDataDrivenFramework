'================================
'Author: Mircea Sirghi, hakkussg@gmail.com, www.linkedin.com/pub/mircea-sirghi/32/6b5/700/
'================================

Option Explicit

'================================
'Description : Action used to 
'Arguments  :  N/A
'================================
Function ActionTemplate()
	Dim fname:fname = "ActionTemplate"
	gLogger.LogFunctionStart(fname)
	gLogger.LogFunctionEnd(fname)
End Function

'================================
'Description : Action used to search a string in GOOGLE
'Arguments  :  sSheetName(String) - reprezents the input data sheet
'================================
Function ActionSearchString(sSheetName)
	Dim fname:fname = "ActionSearchString"
	gLogger.LogFunctionStart(fname)
	
	Dim strSearch:strSearch = DataTable("Text", sSheetName)
	Dim strUrl:strUrl = DataTable("Url", sSheetName)
	Dim strBrowserType:strBrowserType = DataTable("BrowserType", sSheetName)
	Dim repoBrowser:repoBrowser = getBrowser(strBrowserType)
	
	openBrowser(strBrowserType)	
		
	Browser(repoBrowser).BrowserMaximize()	
		
	Browser(repoBrowser).Navigate strUrl
	
	Browser(repoBrowser).Page("Google").WebEdit("txtSearchField").fnSetValue(strSearch)
	
	Browser(repoBrowser).WaitForPageToLoad()
	
	If(Browser(repoBrowser).Page("Google").WebElement("btnStartupSearch").Exist(1)) Then
		Browser(repoBrowser).Page("Google").WebElement("btnStartupSearch").fnClickSecure
	Else
		Browser(repoBrowser).Page("Google").WebElement("btnSearch").fnClickSecure
	End If
	
	Browser(repoBrowser).WaitForPageToLoad()
	
	gLogger.LogFunctionEnd(fname)
End Function
'================================