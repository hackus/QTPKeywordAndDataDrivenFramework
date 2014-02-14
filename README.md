QTPKeywordAndDataDrivenFramework
================================

This framework exemplifies how can QTP be used in a style similar to cucumber.

Contact details:<br>
mail: hakkussg@gmail.com<br>
<a href="http://www.linkedin.com/pub/mircea-sirghi/32/6b5/700/" target="_blank">Find me on linkedin.</a>

One of the most important function in website automation is the WaitForPageToLoad which somehow is not available in QTP. Here is how I am doing this and it works: 
<pre>
<code>
'================================
'Description : Function used to execute javascript code in QTP
'Arguments   : obrowser(Browser)
'            : sJavaScript(String) javascript code to be executed
'Source      : http://www.softwareinquisition.com/81.htm
'================================
Public Function evalJS(oBrowser, sJavaScript, bSeverity)
	Dim fname:fname = "evalJS"	
	'gLogger.LogFunctionStart(fname)
	
	evalJS = ""
	Dim JSEntry
	Set JSEntry = oBrowser.object.document.documentelement.parentnode.parentwindow
	
	Err.Number = 0
	
	On Error Resume Next
	evalJS = JSEntry.eval(sJavaScript)
	
	'Call gLogger.LogWrite("evalJs Value ["&evalJS&"]")
	If(Err.Number <> 0 And bSeverity = true)Then
		Call gLogger.LogException(fname, "Error occured while running the javascript code : [" & sJavaScript & "]")
	End IF

	On Error Goto 0
	
	'gLogger.LogFunctionEnd(fname)
End Function
'================================

'================================
'Description : Function used to make QTP wait till the page content is loaded
'Arguments   : obrowser(Browser)
'Author: Mircea Sirghi(hakkussg@gmail.com, <a href="http://www.linkedin.com/pub/mircea-sirghi/32/6b5/700/" target="_blank">Find me on linkedin.</a> )
'================================
Function WaitForPageToLoad(oBrowser)
	Dim fname:fname = "evalJS"	
	gLogger.LogFunctionStart(fname)
	Dim i:i=0
	Dim completed:completed=false
	
	Dim javaScriptCode
	
	javaScriptCode = " function test()" _
				& " {" _
				& " return document.readyState;" _ 
				& " }" _
				& " test();"
	Do 
		Wait 1
		i=i+1
		If((StrComp(evalJS(oBrowser, javaScriptCode, false),"complete") = 0) OR (StrComp(evalJS(oBrowser, javaScriptCode, false),"interactive") = 0))Then
			completed = true
			Exit Do
		End If
	Loop While i<30
	
	Dim testVal:testVal = evalJS(oBrowser, javaScriptCode, true)
	
	gLogger.LogWrite("evaluation complete value is [" & testVal & "]")
	
	
	If(completed <> true)Then
		Call gLogger.LogWarning(fname, "The page completed state is not 'complete' it might be that the test have run while the page was not completely loaded.")
	End If	
	gLogger.LogFunctionEnd(fname)
End Function
'================================
</code>
</pre>
