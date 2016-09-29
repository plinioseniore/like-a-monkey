' *******************************************************************************************************
' Macro Automatically Executed at opening of the Excel file
' 
' Use the name of the Excel file as key for formulas inside the file, then save the
' contents as value to allow copy/paste into other programs (like Autocad) *******************************************************************************************************
'
Sub auto_open()

If ThisWorkbook.Name <> "DCS_IO_Template.xlsm" Then

    ' Get Filename and cut the extention
    Cells(41, 7).Value = Left(ThisWorkbook.Name, 13)

    'Force formula
    ActiveSheet.EnableCalculation = False
    ActiveSheet.EnableCalculation = True

    'Paste values
    Cells.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

    ' Select print area and copy
    Application.Goto Reference:="Print_Area"
    Selection.Copy
    ThisWorkbook.Save
End If

End Sub