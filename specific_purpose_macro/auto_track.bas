' *******************************************************************************************************
' This macro is used to track the state of a test marked down as a color on a Tagname, it runs
' automatically once the Excel file is saved and build a new file saved as CSV
'
' Tagname, Test, TestState, ColorMark, Date, Tester
' tag1,FileName.xlsm,Passed,50B000,20161115,Plinio
' tag2,FileName.xlsm,Passed,50B000,20161115,Plinio
' tag3,FileName.xlsm,Passed,50B000,20161115,Plinio
' tag4,FileName.xlsm,Failed,1159C6,20161115,Plinio
' tag5,FileName.xlsm,Failed,1159C6,20161115,Plinio
' tag6,FileName.xlsm,ToDo,FFFFFF,20161115,Plinio
' tag7,FileName.xlsm,ToDo,FFFFFF,20161115,Plinio
'
' Place this macro in ThisWorkbook, otherwise will not have effect
' *******************************************************************************************************
'
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)

' Use this values to match the column with the tagnames, this colum shall be marked
' with Red and Green color based on the test state

Colmn = 2       ' Column
LastRow = 50    ' Number of rows to parse

' Build the CSV file
CellData = vbCrLf + "Tagname, Test, TestState, ColorMark, Date, Tester" + vbCrLf  					' Header
FilePath = ThisWorkbook.Path + "\" + ThisWorkbook.Name + "_r" + Format(Now(), "yyyymmdd") + ".txt"  ' Filename

' Parse all Tagnames
For i = 1 To LastRow
    If Cells(i, Colmn) <> "" Then
    
        TestState = "ToDo"  ' If is not marked is to do
    
        ' Get the color
        sColor = Right("000000" & Hex(Cells(i, Colmn).Interior.Color), 6)
        R = Right(sColor, 2)
        G = Mid(sColor, 3, 2)
        B = Left(sColor, 2)
        
        ' Identify the dominant color: Red for Failed and Green for Passed
        If sColor = "FFFFFF" Then
            TestState = "ToDo"
        ElseIf R > G Then
            TestState = "Failed"
        ElseIf G > R Then
            TestState = "Passed"
        End If
    
        ' Build the row
        CellData = CellData + Cells(i, Colmn) + "," + ThisWorkbook.Name + "," + TestState + "," + Hex(Cells(i, Colmn).Interior.Color) + "," + Format(Now(), "yyyymmdd") + "," + Environ("COMPUTERNAME") + "\" + Environ("USERNAME") + vbCrLf
    End If
Next i

' Save the file (will overwrite if exist)
Open FilePath For Output As #1
Write #1, CellData
Close #1

End Sub
