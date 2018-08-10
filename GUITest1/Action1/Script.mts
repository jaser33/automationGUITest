'Open the Browser

Dim iURL 
Dim objShell

'Navigate to the Web App
iURL = "https://webidetesting9467895-p2000517592trial.dispatcher.hanatrial.ondemand.com/webapp/extended_runnable_file.html?hc_orionpath=%2Fp2000517592trial%24P2000517592-OrionContent%2FManageProducts&origional-url=index.html&sap-ui-appCacheBuster=..%2F..%2F&sap-ui-xx-componentPreload=off&sap-language=EN"

set objShell = CreateObject("Shell.Application")
objShell.ShellExecute "iexplore.exe", iURL, "", "", 1

'Start Test 1

'Click the Cheap link
Browser("Unit tests for ManageProduct").Sync

If Browser("Unit tests for ManageProduct").Page("ManageProducts").SAPUITabStrip("SAPUITabStrip").Exist Then
	Browser("Unit tests for ManageProduct").Page("ManageProducts").SAPUITabStrip("SAPUITabStrip").Select "Cheap"
	Wait(2)
Else
	Wait(30)
	Browser("Unit tests for ManageProduct").Page("ManageProducts").SAPUITabStrip("SAPUITabStrip").Select "Cheap"
	Wait(2)
End If


id1 = Browser("Unit tests for ManageProduct").Page("ManageProducts").WebElement("__text2-__clone104").GetROProperty("innertext")

id2 = Browser("Unit tests for ManageProduct").Page("ManageProducts").Link("__identifier0-__clone104-link").GetROProperty("innertext")


Set RegEx = CreateObject("vbscript.regexp") 
RegEx.Pattern = "[^\d]"
RegEx.IgnoreCase = True 
RegEx.Global = True

numStr=RegEx.Replace(id1, "") 
id1a = numStr

numStr2=RegEx.Replace(id2, "") 
id2a = numStr2


If id1a = id2a Then
	Reporter.ReportEvent micPass, "Product ID matches", "Step Passed"
Else
	Reporter.ReportEvent micFail, "Product ID does not match", "Step Failed"
	ExitTest
End If


'Click the Product on the top of the list
Browser("Unit tests for ManageProduct").Page("ManageProducts").SAPUITable("Products (2)Search").SelectItemInCell 1,"Product","ProductID 61"
wait(2)

'Verify that the product width is "4881.75
d1 = Browser("Unit tests for ManageProduct").Page("ManageProducts").WebElement("__text27").GetROProperty("innertext")

If inStr(d1, "4881.75") Then
	Reporter.ReportEvent micPass, "Width is correct", "Step Passed"
Else
	Reporter.ReportEvent micFail, "Width is not correct", "Step Failed"
	ExitTest
End If


'
'
''Set up the arrayList
'Dim alMappings : Set alMappings = CreateObject("System.Collections.ArrayList")
'
'Dim objNetwork : Set objNetwork = WScript.CreateObject("WScript.Network")
'Dim oShell     : Set oShell = CreateObject("WScript.Shell")
'
''## Mapping Objects ##
'
''Position one mapping
'If numRegs >= 1 Then
'	Dim udtPosOne : Set udtPosOne = New Mapping
'	With udtPosOne
'		.strLocalDrive = "P:"
'		.strUNCPath = "\\10.0." & storeNumber & ".101\dpalm"
'		.strPersistent = "False"
'		.strUsr = "somename"
'		.strPas = "somepass"
'	End With
'	alMappings.add(udtPosOne)
'End If
'
''Use i for loop counters
'For i = 0; i < 100; i++	
'	'Use j for the nested loop counters
'	For j = 0; j < 100; j++		
'	Next
'Next
'
'Function fnSum_Num(varOne,varTwo)
''Function fnSum_Num returns the sum of two numbers
'	sumOfNumbers = varOne + varTwo
'	msgbox "The Sum of Numbers is: " & sumOfNumbers
'End Function
'
'

