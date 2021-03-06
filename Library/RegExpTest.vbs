'Function library Number: RegExp CheckString
'Test infomation: Check String Matching string via RegExpression 
'Input value:
'patrn:The searched values
' myString:the prarent string need to be searched
'Output value:
'RegExpTest: the result
'Created by: Ashley.shen
'Created date:8/24/2012

Function RegExpTest(patrn, myString)
	 Dim regEx, Match, Matches      ' Create variable.
	Set regEx = New RegExp         ' Create a regular expression.
	regEx.Pattern = patrn         ' Set pattern.
	regEx.IgnoreCase = True         ' Set case insensitivity.
	regEx.Global = True         ' Set global applicability.
	Dim RetStr
	If myString <> "" Then
				If (regEx.Test(myString) = True) Then
						Set Matches = regEx.Execute(myString)   ' Execute search.
								For Each Match in Matches      ' Iterate Matches collection.
									RetStr = RetStr & "Match found at position "
									RetStr = RetStr & Match.FirstIndex & ". Match Value is '"
									RetStr = RetStr & Match.Value & "'." & vbCRLF
								Next
						Reporter.ReportEvent micDone,"Verify Expected String in desired String or not","Expected String has been found in the Desired String" & "RetStr is:" & RetStr
				Else
					RetStr = "String Matching Failed"
					Reporter.ReportEvent micDone,"Verify Expected String in desired String or not","Expected String did not be found in the Desired String" & "RetStr is:" & RetStr
			   End If
			RegExpTest = RetStr
	Else
	RetStr = "Desired String is Empty"
	Reporter.ReportEvent micFail,"Verify Desired String is Empty or not","The desired String is Empty" & "RetStr is:" & RetStr
	End If
End Function
