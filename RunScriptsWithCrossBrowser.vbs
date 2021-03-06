Dim arrayReplaceBrowserInfo
arrayReplaceBrowserInfo=array("<Name>targetBrowserIndex</Name>		<Value>0</Value>","<Name>targetBrowserIndex</Name>		<Value>1</Value>","<Name>targetBrowserIndex</Name>		<Value>2</Value>")
intArrayLength= ubound(arrayReplaceBrowserInfo)-lbound(arrayReplaceBrowserInfo)+1

Dim testEnvironmentConfigurationFilePath
testEnvironmentConfigurationFilePath=".\TestData\TestEnvironmentConfiguration.xml"
Dim reportLocationPath
reportLocationPath="C:\Users\serena.jin\.hudson\jobs\EcommerceTesting-CrossBrowser\workspace\TestReport"

For i=0 to intArrayLength-1.
	Call killProcess("MultiTestManager.exe")
	updateTestEnvironmentConfigurationFile arrayReplaceBrowserInfo(i),testEnvironmentConfigurationFilePath
	doBatchRun(reportLocationPath)
Next


Function returnMatch(searchText, targetString )
   	Dim regExp  ' Create variable.
	Set regExp = New RegExp         ' Create a regular expression.
	regExp.Pattern = searchText         ' Set pattern.
	regExp.IgnoreCase = True         ' Set case insensitivity.
	regExp.Global = True         ' Set global applicability.
   Set Matches = 	regExp.Execute(targetString)
   returnMatch 	= 	Matches(0)
End Function

Function doBatchRun(reportLocation)
	'Create an Instance of Multi Test Manager
	Set oMTM = CreateObject("MultiTestManager.Application")
	oMTM.Visible = True
	
'	dir = "C:\Users\serena.jin\.hudson\jobs\EcommerceTesting-CrossBrowser\workspace"
	dir ="."
	'Change Default Report Settings
	Set oReportSettings = oMTM.Preferences.ReportSettings
	oReportSettings.CreateReport = True
	oReportSettings.OverwriteReport = False
	oReportSettings.DefaultLocation = False
	'oReportSettings.ReportLocation =  "C:\Users\serena.jin\.hudson\jobs\EcommerceTesting-CrossBrowser\workspace\TestReport"
	oReportSettings.ReportLocation = reportLocation
	oReportSettings.ReportName = "Ecommerce Cross Browser Test Report"
	oReportSettings.ViewReport = False
	oReportSettings.AddCustomStep "Run by", "Hudson", True
	
	'Add all ecommerce test scripts with name started by "CAT_" to batch run in Multi Test Manager
	Set fso=createobject("scripting.filesystemobject")
	set parentFolder=fso.getfolder(dir&"\TestScripts\")
	'Add the test script "GDC_Server_Connect_VPN"
'	oMTM.AddTestScript parentFolder.Path &"\" &"GDC_Server_Connect_VPN", True 
	For each subFolder in parentFolder.SubFolders
		Set regEx = New RegExp         
		regEx.Pattern = "CAT_"         ' Set pattern.
		regEx.IgnoreCase = True         ' Set case insensitivity.
		regEx.Global = True         ' Set global applicability.
		If (regEx.Test(subFolder.Name) = True) Then
			oMTM.AddTestScript parentFolder.Path &"\" &subFolder.Name, True
		End If
	Next
	'Add the test script "GDC_Server_Disconnect_VPN"
'	oMTM.AddTestScript parentFolder.Path &"\" &"GDC_Server_Disconnect_VPN", True
	
	'Run the scripts by Mutlti Test Manager
	oMTM.Run
	
	while ( oMTM.IsRunning )
	Wend
	
	Set oRunSettingss = Nothing
	Set oReportSettings = Nothing
	
	'Close Multi Test Manager
	oMTM.Quit
	Set oMTM = Nothing
End Function 

Function updateTestEnvironmentConfigurationFile(browserInfo,filePath)
	Set fso=Wscript.createobject("scripting.filesystemobject")	
	Set updatedFile=fso.OpenTextFile(filePath,1,true)
	allFileInfo=updatedFile.ReadAll

'msgbox "AllFileInfo:"&allFileInfo
	matchedValue = returnMatch( "<Name>targetBrowserIndex</Name>[\s\n]*<Value>\d*</Value>", allFileInfo)
	allFileInfo=replace(allFileInfo,matchedValue,browserInfo)
	updatedFile.Close
	Set r=fso.OpenTextFile(filePath,2,true)
	r.Write allFileInfo
	r.Close
	set updatedFile =nothing
End Function

Function moveDetailedReport()
Set fso=createobject("scripting.filesystemobject")
Set workspacefolder = fso.getfolder(dir)
buildParentFolderPath = workspacefolder.ParentFolder.Path
set parentTestFolder= fso.getfolder(buildParentFolderPath&"\builds")

For each folder1 in parentTestFolder.SubFolders
   detailedReportParentPath =  parentTestFolder.Path & "\" &folder1.Name
   Set detailedReportParentFolder = fso.GetFolder(detailedReportParentPath)
   For each folder2 in detailedReportParentFolder.SubFolders
           If   fso.FolderExists(detailedReportParentPath  &"\htmlreports\Ecommerce_Summary_Report") and  fso.FolderExists( detailedReportParentPath & "\archive")Then
			   	fso.CopyFolder  detailedReportParentPath & "\archive\*", detailedReportParentPath  &"\htmlreports\Ecommerce_Summary_Report", true
		   End If
   Next
Next
Set fso=Nothing
End Function


Function killProcess(strProcessToKill)
Const strComputer = "." 
 Dim objWMIService, colProcessList
 Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
 Set colProcessList = objWMIService.ExecQuery("SELECT * FROM Win32_Process WHERE Name ='" & strProcessToKill & "'")
 For Each objProcess in colProcessList 
objProcess.Terminate() 
 Next
End Function



