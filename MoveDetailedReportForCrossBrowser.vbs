dir = "."
'Move detailed reports to foloder Ecommerce_Summary_Report 
Set fso=createobject("scripting.filesystemobject")
Set workspacefolder = fso.getfolder(dir)
buildParentFolderPath = workspacefolder.ParentFolder.Path
set parentTestFolder= fso.getfolder(buildParentFolderPath&"\builds")
For each folder1 in parentTestFolder.SubFolders
   detailedReportParentPath =  parentTestFolder.Path & "\" &folder1.Name
   Set detailedReportParentFolder = fso.GetFolder(detailedReportParentPath)
   For each folder2 in detailedReportParentFolder.SubFolders
		   If   fso.FolderExists(detailedReportParentPath  &"\htmlreports\Ecommerce_Summary_Report_For_IE") and  fso.FolderExists( detailedReportParentPath & "\archive")Then
			   	fso.CopyFolder  detailedReportParentPath & "\archive\*", detailedReportParentPath  &"\htmlreports\Ecommerce_Summary_Report_For_IE", true
			end if
			If fso.FolderExists(detailedReportParentPath  &"\htmlreports\Ecommerce_Summary_Report_For_Firefox") and  fso.FolderExists( detailedReportParentPath & "\archive")Then
			   	fso.CopyFolder  detailedReportParentPath & "\archive\*", detailedReportParentPath  &"\htmlreports\Ecommerce_Summary_Report_For_Firefox", true
			end if
			If fso.FolderExists(detailedReportParentPath  &"\htmlreports\Ecommerce_Summary_Report_For_Chrome") and  fso.FolderExists( detailedReportParentPath & "\archive")Then
			   	fso.CopyFolder  detailedReportParentPath & "\archive\*", detailedReportParentPath  &"\htmlreports\Ecommerce_Summary_Report_For_Chrome", true
		   End If
   Next
Next
Set fso=Nothing


