Public Sub ExelToWord()

Application.ScreenUpdating = False

' First row where the macro look for certificates name
i = 2

' Parse all Certificates
Do While Worksheets("Data").Cells(i, 1).Value <> ""

    ' Select Worksheet
    Worksheets("INPUTS").Activate

    ' Read data
    Certificate = Worksheets("Data").Cells(i, 1).Value
    Dato1 = Worksheets("Data").Cells(i, 2).Value
    Dato2 = Worksheets("Data").Cells(i, 3).Value
    Dato3 = Worksheets("Data").Cells(i, 4).Value
    Dato4 = Worksheets("Data").Cells(i, 5).Value
    Dato5 = Worksheets("Data").Cells(i, 6).Value
    Dato6 = Worksheets("Data").Cells(i, 7).Value
    Dato7 = Worksheets("Data").Cells(i, 8).Value
   
    ' Write data in the INPUT sheet
    ActiveSheet.Cells(2, 9).Value = Dato1
    ActiveSheet.Cells(4, 9).Value = Dato2
    ActiveSheet.Cells(6, 9).Value = Dato3
    ActiveSheet.Cells(8, 9).Value = Dato4
    ActiveSheet.Cells(10, 9).Value = Dato5
    ActiveSheet.Cells(19, 9).Value = Dato6
    ActiveSheet.Cells(58, 9).Value = Dato7
    
    
    ' Select Worksheet
    Worksheets("ENGLISH").Activate
      
    ' Select print area and copy, it works if the target worksheet has a PrintArea selected
    Application.Goto Reference:="Print_Area"
    Selection.Copy

    ' Create a Word document
    Dim objWord
    Dim objDoc
    Dim objSelection
    Set objWord = CreateObject("Word.Application")
	' Use a template DOTX to format the Word document
    Set objDoc = objWord.Documents.Add(Template:=ActiveWorkbook.Path & "\gaztemplate.dotx")
        objWord.Visible = True
    Set objSelection = objWord.Selection
    
    ' Paste data from Excel to Word
    objSelection.Paste
    objSelection.InsertBreak
        
    ' Select Worksheet
    Worksheets("RUSSIAN").Activate
      
    ' Select print area and copy, it works if the target worksheet has a printarea selected
    Application.Goto Reference:="Print_Area"
    Selection.Copy
    
    ' Paste data from Excel to Word
    objSelection.Paste

    ' Save and Close
    objDoc.SaveAs2 (ActiveWorkbook.Path & "\" & Certificate & ".docx")
    objDoc.Close
    objWord.Quit
    
    ' Next one
    i = i + 1

Loop

Application.ScreenUpdating = True

End Sub


