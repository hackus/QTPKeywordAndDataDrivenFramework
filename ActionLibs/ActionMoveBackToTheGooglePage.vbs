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
'Description : Action used to move browser back to the GOOGLE home page
'Arguments  :  N/A
'================================
Function ActionMoveBackToTheGooglePage(sSheetName)
	Dim fname:fname = "ActionMoveBackToTheGooglePage"
	gLogger.LogFunctionStart(fname)
	
	Dim strBrowserType:strBrowserType = DataTable("BrowserType", sSheetName)
	Dim repoBrowser:repoBrowser = getBrowser(strBrowserType)
	Dim i:i=0
	Do
		i=i+1
		Browser(repoBrowser).Back
		Browser(repoBrowser).WaitForPageToLoad()	
		
		If (Browser(repoBrowser).Page("Google").WebEdit("txtSearchField").Exist(1)) Then
			Exit Do
		End If
	Loop While i<30

	gLogger.LogFunctionEnd(fname)
End Function
'================================