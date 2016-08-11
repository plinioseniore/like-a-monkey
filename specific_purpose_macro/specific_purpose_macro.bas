Sub Export_Process_JB()
'
' This macro export the tags and relevant JB
'

nRow = 0    ' Row Index
nWires = 0  ' Number of wires
nBlanks = 0 ' Number of blank rows

Tagname = ""

' Go through all the worksheets
nSheet = Worksheets.Count

' Start from first tab that contains a JB
For i = 9 To nSheet
    
    ' Go through all the cells
    For j = 17 To 200

        ' Get the tagname
        If (Worksheets(i).Cells(j, 1).Value <> "") Then
            
            ' Save tagname
            Tagname = Worksheets(i).Cells(j, 1).Value
            
            ' If we are here there are no more signal to process
            If (Tagname = "Note 1") Then Exit For
            
            ' Next row
            nRow = nRow + 1
            nBlanks = 0

            ' Fill values
            Cells(nRow, 1).Value = Worksheets(i).Name  ' JB Name
            Cells(nRow, 2).Value = Worksheets(i).Range("E9").Value  ' JB Name
            Cells(nRow, 3).Value = Tagname                          ' Tagname
            Cells(nRow, 4).Value = Worksheets(i).Cells(j, 4).Value  ' Wire color
            
            ' Add color from other wires (up to 6)
            For k = 1 To 6
                If (Worksheets(i).Cells(j + k, 1).Value = "") Then
                    Cells(nRow, 4 + k).Value = Worksheets(i).Cells(j + k, 4).Value  ' Wire color in column 4
                Else
                    Exit For
                End If
            Next k
        
        Else
            
            ' If there are too many white spaces, we are at end of JB
            If (nBlanks > 10) Then Exit For
        
            nBlanks = nBlanks + 1
        
        End If
        
    Next j

Next i
'
End Sub

' Parse all the cable and assign a module with a constrain on the number
' of maximum signal for each module
'
' Need a list of all cables and a list of all modules
'
Sub CableAssignment()

' Go through all the cables
For i = 1 To 300

    ' If there are no more cables to assign
    If (Worksheets(1).Cells(i, 4) = "") Then Exit For
    
    ' Go though the modules
    For k = 1 To 100
    
        ' If there are no more modules to assign
        If (Worksheets(2).Cells(k, 3) = "") Then Exit For
    
        ' Match module type and area
        If (Worksheets(1).Cells(i, 3) = Worksheets(2).Cells(k, 1)) Then
        
            ' Verify if there are enough channels
            If (Worksheets(2).Cells(k, 5) + Worksheets(1).Cells(i, 5) <= Worksheets(2).Cells(k, 4)) Then
            
                ' If we are here, there is room on the module
                Worksheets(2).Cells(k, 5) = Worksheets(2).Cells(k, 5) + Worksheets(1).Cells(i, 5)
                Worksheets(1).Cells(i, 6) = Worksheets(2).Cells(k, 3)
                
                ' Next cable
                Exit For
            End If
            
        
        End If
    
    
    Next k
    
Next i

End Sub


Sub Export_ELE()
'
' This macro export the tags and relevant JB
'

nRow = 0    ' Row Index
nWires = 0  ' Number of wires
nBlanks = 0 ' Number of blank rows

Tagname = ""

' Go through all the worksheets
nSheet = Worksheets.Count

' Start from first tab that contains a JB
For i = 2 To nSheet
    
    ' Go through all the cells
    For j = 5 To 200

        ' Get the tagname
        If (Worksheets(i).Cells(j, 1).Value <> "") Then
            
            ' Save tagname
            Tagname = Worksheets(i).Cells(j, 1).Value
            
            ' If we are here there are no more signal to process
            If (Tagname = "Note 1") Then Exit For
            
            ' Next row
            nRow = nRow + 1
            nBlanks = 0

            ' Fill values
            Cells(nRow, 1).Value = Worksheets(i).Name  ' JB Name
            Cells(nRow, 2).Value = Worksheets(i).Range("E5").Value  ' JB Name
            Cells(nRow, 3).Value = Tagname                          ' Tagname
            Cells(nRow, 4).Value = Worksheets(i).Cells(j, 4).Value  ' Wire color
            Cells(nRow, 10).Value = Worksheets(i).Cells(j, 2).Value  ' Terminal number
            
            
            ' Add color from other wires (up to 6)
            For k = 1 To 6
                If (Worksheets(i).Cells(j + k, 1).Value = "") Then
                    Cells(nRow, 4 + k).Value = Worksheets(i).Cells(j + k, 4).Value  ' Wire color in column 4
                    Cells(nRow, 10 + k).Value = Worksheets(i).Cells(j + k, 2).Value ' Terminal number
                Else
                    Exit For
                End If
            Next k
        
        Else
            
            ' If there are too many white spaces, we are at end of JB
            If (nBlanks > 10) Then Exit For
        
            nBlanks = nBlanks + 1
        
        End If
        
    Next j

Next i


'
End Sub


Sub Export_CE()
'
' This macro export the tags and relevant JB
'

nRow = 0    ' Row Index
nWires = 0  ' Number of wires
nBlanks = 0 ' Number of blank rows

Tagname = ""

' Go through all the worksheets
nSheet = Worksheets.Count

' Start from first tab that contains a JB
For i = 3 To nSheet
    
    ' Go through all the causes
    For j = 32 To 200

        ' Get the tagname
        If (Worksheets(i).Cells(j, 8).Value <> "") Then
            
            ' Save tagname
            Tagname = Worksheets(i).Cells(j, 8).Value
            
            ' Next row
            nRow = nRow + 1
            nBlanks = 0

            ' Fill values
            Cells(nRow, 1).Value = Worksheets(i).Name                ' CE Name
            Cells(nRow, 2).Value = Worksheets(i).Range("F10").Value  ' CE Name
            Cells(nRow, 3).Value = Tagname                          ' Tagname
            Cells(nRow, 4).Value = Worksheets(i).Cells(j, 5).Value  ' Service
        
        Else
            
            ' If there are too many white spaces, we are at end of JB
            If (nBlanks > 30) Then Exit For
        
            nBlanks = nBlanks + 1
        
        End If
        
    Next j
    
     nBlanks = 0
    
     ' Go through all the causes
    For j = 12 To 65

        ' Get the tagname
        If (Worksheets(i).Cells(28, j).Value <> "") Then
            
            ' Save tagname
            Tagname = Worksheets(i).Cells(28, j).Value
            
            ' Next row
            nRow = nRow + 1
            nBlanks = 0

            ' Fill values
            Cells(nRow, 1).Value = Worksheets(i).Name                ' CE Name
            Cells(nRow, 2).Value = Worksheets(i).Range("F10").Value  ' CE Name
            Cells(nRow, 3).Value = Tagname                          ' Tagname
            Cells(nRow, 4).Value = Worksheets(i).Cells(10, j).Value  ' Service
        
        Else
            
            ' If there are too many white spaces, we are at end of JB
            If (nBlanks > 30) Then Exit For
        
            nBlanks = nBlanks + 1
        
        End If
        
    Next j
    

Next i
'
End Sub


Sub LoopVSTag()

Loop = ""
Column = 0

' Move though all loops
For i = 2 To 1167
	
	Column = 0
	Loop = Cells(i, 1).Value
	
	' Move though all tags
	For j = 2 To 3729
		If (Loop = Worksheets(1).Cells(j, 14).Value)
			Cells(i, 2+Column).Value = Worksheets(1).Cells(j, 5).Value
		End If
	Next j
Next i

End Sub
