	'Create an Instance of Multi Test Manager
	Set oMTM = CreateObject("MultiTestManager.Application")
	oMTM.Visible = True
	
	dir ="."
	'Change Default Report Settings
	Set oReportSettings = oMTM.Preferences.ReportSettings
	oReportSettings.CreateReport = True
	oReportSettings.OverwriteReport = False
	oReportSettings.DefaultLocation = True
	oReportSettings.ReportName = "Ecommerce Cross Browser Test Report"
	oReportSettings.ViewReport = True
	oReportSettings.AddCustomStep "Run by", "Hudson", True
	
	'Add all ecommerce test scripts with name started by "CAT_" to batch run in Multi Test Manager
	Set fso=createobject("scripting.filesystemobject")
	set parentFolder=fso.getfolder(dir&"\TestScripts\")
	For each subFolder in parentFolder.SubFolders
		Set regEx = New RegExp         
		regEx.Pattern = "CAT_"         ' Set pattern.
		regEx.IgnoreCase = True         ' Set case insensitivity.
		regEx.Global = True         ' Set global applicability.
		If (regEx.Test(subFolder.Name) = True) Then
			oMTM.AddTestScript parentFolder.Path &"\" &subFolder.Name, True
		End If
	Next


