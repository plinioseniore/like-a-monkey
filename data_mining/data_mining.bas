' *******************************************************************************************************
' Look for a matching Target value into the CellRange that match the LoopNumber. The first column in
' CellRange is used to match the LoopNumber rather the last to match the Target.
'
' The function output is the number of times that Target is found into CellRange for the LoopNumber
'
' |   A      |    B     |  C  |
' |Tagname   |LoopNumber|Type |
' |PT-941001 |P-941001  | AIR |
' |PV-941001A|P-941001  | AOR |
' |PV-941001B|P-941001  | AOR |
'
' PropertyTypeCount("P-941001", B:C, "AO") returns 2.
' PropertyTypeCount("P-941001", B:C, "AI") returns 1.
'
' The optional values Alphabetical and StartingRow allow a faster parsing of the data, assuming that data
' are in alphabetical order and assigning the starting row.
'
' *******************************************************************************************************
'
Public Function PropertyTypeCount(LoopNumber As String, CellRange As Range, Target As String, Optional ByVal Alphabetical As Integer = 0, Optional ByVal StartingRow As Integer = 1) As Variant

' The last row is identified going through the column down to the first empty cell
' use instead NumberOfRows = CellRange.Rows.Count if you don't want to rely on empty cells
NumberOfRows = CellRange.Rows.End(xlDown).Row
NumberOfColumns = CellRange.Columns.Count
TargetNo = 0
Catched = 0

For i = StartingRow To NumberOfRows
    If CellRange.Cells(i, 1) = LoopNumber Then   ' Match the LoopNumber in the first column
        
		' Mark that we have found the LoopNumber, so we can exit at next iteration
		If Alphabetical = 1 Then                      
			Catched = 1
		End If
		
		If InStr(CellRange.Cells(i, NumberOfColumns), Target) >= 1 Then
            TargetNo = TargetNo + 1
        End If
    ElseIf Catched = 1 Then
        Exit For                                 ' Once we passed over, we can exit if LoopName is in alphabetical order
    End If
Next i

PropertyTypeCount = TargetNo

End Function


' *******************************************************************************************************
' Extract the LoopNumer and LoopName from a TagName.
' 
' LoopNumber("PT-941001", FALSE) returns 941001
' LoopNumber("PT-941001", TRUE)  returns P941001
' *******************************************************************************************************
'
Public Function LoopNumber(TagName As String, WithLoopType As Boolean) As String

Dim Result As String
Result = ""

' Parse tha TagName looking for numerical values
For i = 1 To Len(TagName)

	' Once the first numerical has been found, go ahead till the last numerical
    If IsNumeric(Mid(TagName, i, 1)) Then
        c = i
        Do While IsNumeric(Mid(TagName, c, 1))
            Result = Result & Mid(TagName, c, 1)
            c = c + 1
        Loop
            
        Exit For
    End If
Next i

' Add the first letter of the TagName as loop type
If WithLoopType Then
		LoopNumber = Mid(TagName, 1, 1) & Result
Else: 	LoopNumber = Result
End If

End Function

' *******************************************************************************************************
' Search and returns a TagNumber into a string
' 
' Look for a tagnumber in the form of LLLLLUUNNNN
' where:
' LLLL is the tag type (FT, FI, FAL, FALL, ...)
' UU   is the unit number
' NNNN is the loop number
' 
' RetrieveTagNumber("IN DIFFERENT CARD THAN FT941001 TRANSMITTER", "94") returns FT941001
'
' *******************************************************************************************************
'
Public Function RetrieveTagNumber(StringValue As String, UnitNumber As String) As String

For i = 1 To Len(StringValue)

    ' Detect the UU into the tagname
    If Mid(StringValue, i, Len(UnitNumber)) = UnitNumber Then
    
        ' The UU is followed by the loop number, assume at least two digits
        If IsNumeric(Mid(StringValue, i + 1, Len(UnitNumber))) And IsNumeric(Mid(StringValue, i + 2, Len(UnitNumber))) Then
        
           ' The UU has before at tag type, assume at least two digits
            If (i > 2) And (Not IsNumeric(Mid(StringValue, i - 1, Len(UnitNumber)))) And (Not IsNumeric(Mid(StringValue, i - 2, Len(UnitNumber)))) Then
                
                i_before = 0
                i_after = 0
                
                ' Uau, we got a tagname! Assume that a space before and after is there
                For j = 1 To 10
                    If (Mid(StringValue, i - i_before, 1) <> " ") Then i_before = i_before + 1
                    If (Mid(StringValue, i + i_after, 1) <> " ") Then i_after = i_after + 1
                Next j
                
                Exit For
                
            End If
            
        End If
        
    End If
    
Next i

RetrieveTagNumber = Mid(StringValue, i - i_before + 1, i + i_after - 1)

End Function

' *******************************************************************************************************
' Return if a cell has the strike through formatting set
' *******************************************************************************************************
'
Function Is_Strikethrough(r As Range)
  Is_Strikethrough = r.Font.Strikethrough
End Function