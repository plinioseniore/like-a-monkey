' Highlight differences in a row-driven database patch file build using Git
' it require the primary key in the column number 2 and use the + and - key
' to identify what has been added or deleted

Sub CheckForChanges()

' Get the number of used row
usedrow = 1
Do While (Cells(usedrow, 1).Value <> "")
 usedrow = usedrow + 1
Loop

' Get the number of used column
usedcol = 1
Do While (Cells(1, usedcol).Value <> "")
 usedcol = usedcol + 1
Loop

' Check for data
For currentrow = 1 To usedrow
 
 ' First two colums should be in format -NN;TAGNAME or +NN;TAGNAME
 ' for each changed row the - one should came before the + one.
 Tagname = Cells(currentrow, 2).Value
 IsChanged = Left(Cells(currentrow, 1).Value, 1)
 
 ' Raws that has changed are tagged with "-" by Git
 If IsChanged = "-" Then
   
  ' Looks for changes
  Changes = 1
  Do While ((Cells(currentrow + Changes, 2).Value <> Tagname) And ((currentrow + Changes) < usedrow))
   Changes = Changes + 1
  Loop
  
  If ((currentrow + Changes) <> usedrow) Then
  
   'If we are here, a new line with same tagname has been found
   IsChanged = Left(Cells(currentrow + Changes, 1).Value, 1)
   If IsChanged = "+" Then
   
    ' If we are here, a cell has been changed, mark the first cell
	 Cells(currentrow, 1).Interior.ColorIndex = 3
   
    ' We got a new line, we can compare them
    newrow = currentrow + Changes
   
    ' Move inside the row
    For j = 1 To usedcol
     OldValue = Cells(currentrow, j).Value
     NewValue = Cells(newrow, j).Value
    
     ' Mark the change on the new row
     If (NewValue <> OldValue) Then
      Cells(newrow, j).Interior.ColorIndex = 37
      Cells(1, j).Interior.ColorIndex = 37 ' Mark also the top cell
     End If
    Next j
    
    End If
    
   End If
   
  End If

 Next currentrow

End Sub

