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
'Description : Action used to click on the first found link
'Arguments  :  N/A
'================================
Function ActionSelectFirstFoundLink(sSheetName)
	Dim fname:fname = "ActionSelectFirstFoundLink"
	gLogger.LogFunctionStart(fname)
	
	Dim strBrowserType:strBrowserType = DataTable("BrowserType", sSheetName)
	Dim repoBrowser:repoBrowser = getBrowser(strBrowserType)
	
	Dim obj
	Dim oDesc
	Set oDesc = Description.Create
	oDesc("micclass").value = "WebElement"
	oDesc("html tag").value = "H3"
	oDesc("class").value = "r"
	
	'Find all the Links
	Set obj = Browser(repoBrowser).Page("Google").ChildObjects(oDesc)
	
	obj(0).Link("html tag:=A").Click
	
	Browser(repoBrowser).WaitForPageToLoad()

	gLogger.LogFunctionEnd(fname)
End Function
'================================