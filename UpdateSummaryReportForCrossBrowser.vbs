dir = "."
'Update summary report to include the report details of each test cases
Set fso=createobject("scripting.filesystemobject")
set parentTestFolder=fso.getfolder(dir&"\TestReport\")
For each file1 in parentTestFolder.files
	Dim regEx
	Set regEx = New RegExp         
	regEx.Pattern = "Ecommerce Cross Browser Test Report"         ' Set pattern.
	regEx.IgnoreCase = True         ' Set case insensitivity.
	regEx.Global = True         ' Set global applicability.
	If (regEx.Test(file1.Name) = True) Then
		Set htmlSummaryReport = fso.OpenTextFile (parentTestFolder.Path &"\"&file1.Name, 1)
		strText = htmlSummaryReport.ReadAll
		strText = Replace(strText, "\Report\Results.qtp","\Log\LogFile.html")
		strText = Replace(strText, "C:\Users\serena.jin\.hudson\jobs\EcommerceTesting-CrossBrowser\workspace",".")
		strText = Replace(strText, "\","/")
		htmlSummaryReport.Close
		Set finalSummaryReport = fso.OpenTextFile (parentTestFolder.Path &"\"&file1.Name, 2)
		finalSummaryReport.Write(strText)
		finalSummaryReport.Close
		Set htmlSummaryReport= Nothing
		Set finalSummaryReport =Nothing
	End If
Next
Set fso=Nothing



