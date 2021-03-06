'########################################################
'  Name :    open_Browser_Navigate_To_Base_Page
' Objective           :    Launch new browser and navigate to test base page
' Author                 :     Loya.He
' Date Created   :    10/10/2012
' Date Updated :     12/18/2012
'Updated by       :     Loya He
'########################################################

Function open_Browser_Navigate_To_Base_Page (IE, targetUrl)	
	Set IE =  IE
	IE.visible=true
	IE.Navigate targetUrl
End Function

'########################################################
'  Name :    waitForIEPageLoad
' Objective           :    Wait for IE page is loaded sucessfully 
' Author                 :     Loya.He
' Date Created   :    10/20/2012
' Date Updated :     
'Updated by       :    
'########################################################
Function waitForIEPageLoad(IE)
	wait 6
	While  IE.Busy or  IE.ReadyState <> 4
		wait 3
	Wend
End Function


'########################################################
'  Name :    close_Browser
' Objective           :    Close the target browsers
' Author                 :     Loya.He
' Date Created   :    10/20/2012
' Date Updated :     10/26/2012
'Updated by       :    Loya He
'########################################################
Function  close_Browser(desiredBrowser)
	Dim strBrowsers
	strBrowsers = array("iexplore.exe",  "firefox.exe","chrome.exe")
	'Close the browser which need be run for automation test
	If  desiredBrowser = 0 Then
			SystemUtil.CloseProcessByName strBrowsers(0)
	ElseIf desiredBrowser = 1 Then
		SystemUtil.CloseProcessByName strBrowsers(1)
	ElseIf desiredBrowser = 2 Then
		SystemUtil.CloseProcessByName strBrowsers(2)
	End If
End Function


'########################################################
'  Name :   SearchByRegExp
' Objective           :    Search for first Match, return it
' Author                 :     Loya.He
' Date Created   :    10/20/2012
' Date Updated :     10/30/2012
'Updated by       :    Serena Jin
'########################################################
Function SearchByRegExp(regExpPattern, targetString, matchIndex)
	 Dim regExp, Match, Matches      ' Create variable.
	Set regExp = New RegExp         ' Create a regular expression.
	regExp.Pattern = regExpPattern         ' Set pattern.
	regExp.IgnoreCase = True         ' Set case insensitivity.
	regExp.Global = True         ' Set global applicability.

	Set Matches = regExp.Execute(targetString)   ' Execute search.
	Set Match = Matches(matchIndex)
	RetStr = Match
	SearchByRegExp = RetStr
End Function


'########################################################
'  Name :   Contains
' Objective           :    Verify if the searchText is existing in targetString, 
	'If contains, return True
	'Else return False
' Author                 :     Loya.He
' Date Created   :    10/20/2012
' Date Updated :     10/31/2012
'Updated by       :    Serena Jin
'########################################################
Function Contains(searchText, targetString )
   	 Dim regExp  ' Create variable.
	Set regExp = New RegExp         ' Create a regular expression.
	regExp.Pattern = searchText         ' Set pattern.
	regExp.IgnoreCase = True         ' Set case insensitivity.
	regExp.Global = True         ' Set global applicability.
		If (regExp.Test(targetString) = True) Then
			Contains = True
			Else Contains= False
		End If
End Function


'##############################################################################################################################
'  Name :   ReturnSubMatch
' Objective           :    Get the first Match in targetString using RegExp, The get the subMatch by SubMatchIndex
' When a regular expression is executed, zero or more submatches can result when subexpressions are enclosed in capturing parentheses.
' Each item in the SubMatches collection is the string found and captured by the regular expression.
' Author                 :     Loya.He
' Date Created   :    10/20/2012
' Date Updated :     10/31/2012
'Updated by       :    Serena Jin
'##############################################################################################################################
Function ReturnSubMatch(regExp, targetString, SubMatchIndex)
  Dim oRe, oMatch, oMatches
  Set oRe = New RegExp
  ' Set pattern.
  oRe.Pattern = regExp
  ' Get the Matches collection
  Set oMatches = oRe.Execute(targetString)
  ' Get the first item in the Matches collection
  Set oMatch = oMatches(0)
  ' Get the sub-matched parts of first Match.
  ReturnSubMatch = oMatch.SubMatches(SubMatchIndex)
End Function

'##############################################################################################################################
'  Name :   ReturnSubMatchWithIgnoreCase
' Objective           :    Get the first Match in targetString using RegExp, The get the subMatch by SubMatchIndex
' When a regular expression is executed, zero or more submatches can result when subexpressions are enclosed in capturing parentheses.
' Each item in the SubMatches collection is the string found and captured by the regular expression.
' Author                 :    Loya He
' Date Created   :    10/20/2012
' Date Updated :     
'Updated by       :   
'##############################################################################################################################
Function ReturnSubMatchWithIgnoreCase(regExp, targetString, SubMatchIndex)
  Dim oRe, oMatch, oMatches
  Set oRe = New RegExp
	oRe.IgnoreCase = True
  ' Set pattern.
  oRe.Pattern = regExp
  ' Get the Matches collection
  Set oMatches = oRe.Execute(targetString)
  ' Get the first item in the Matches collection
  Set oMatch = oMatches(0)
  ' Get the sub-matched parts of first Match.
  ReturnSubMatchWithIgnoreCase = oMatch.SubMatches(SubMatchIndex)
End Function


'##############################################################################################################################
'  Name :   get cell value from desired Excel file and exact Row No and Column No
' Objective           :    Get the cell value from excel based Row No and Column No
' Author                 :    Loya He
' Date Created   :    03/16/2014
' Date Updated :     
'Updated by       :   
'##############################################################################################################################
Function getcellValueFrmExcel(filepath,sheetname,rownum,colnum)
	Set excobj = createobject("Excel.Application")
	Set excfile = excobj.workbooks.open(filepath)
	Set excsheet = excfile.worksheets(sheetname)
	cellValue = excsheet.cells(rownum,colnum)
	excfile.Close
End Function

