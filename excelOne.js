// Excel macro to copy and paste cels
Sub Range_Copy_Examples()
'Use the Range.Copy method for a simple copy/paste

    'The Range.Copy Method - Copy & Paste with 1 line
    Range("A1").Copy Range("C1")
    Range("A1:A3").Copy Range("D1:D3")
    Range("A1:A3").Copy Range("D1")
    
    'Range.Copy to other worksheets
    Worksheets("Sheet1").Range("A1").Copy Worksheets("Sheet2").Range("A1")
    
    'Range.Copy to other workbooks
    Workbooks("Book1.xlsx").Worksheets("Sheet1").Range("A1").Copy _
        Workbooks("Book2.xlsx").Worksheets("Sheet1").Range("A1")

End Sub


Sub Paste_Values_Examples()
'Set the cells' values equal to another to paste values

    'Set a cell's value equal to another cell's value
    Range("C1").Value = Range("A1").Value
    Range("D1:D3").Value = Range("A1:A3").Value
     
    'Set values between worksheets
    Worksheets("Sheet2").Range("A1").Value = Worksheets("Sheet1").Range("A1").Value
     
    'Set values between workbooks
    Workbooks("Book2.xlsx").Worksheets("Sheet1").Range("A1").Value = _
        Workbooks("Book1.xlsx").Worksheets("Sheet1").Range("A1").Value
        
End Sub


Sub PasteSpecial_Examples()
'Use the Range.PasteSpecial method for other paste types

    'Copy and PasteSpecial a Range
    Range("A1").Copy
    Range("A3").PasteSpecial Paste:=xlPasteFormats
    
    'Copy and PasteSpecial a between worksheets
    Worksheets("Sheet1").Range("A2").Copy 
    Worksheets("Sheet2").Range("A2").PasteSpecial Paste:=xlPasteFormulas
    
    'Copy and PasteSpecial between workbooks
    Workbooks("Book1.xlsx").Worksheets("Sheet1").Range("A1").Copy
    Workbooks("Book2.xlsx").Worksheets("Sheet1").Range("A1").PasteSpecial Paste:=xlPasteFormats
    
    'Disable marching ants around copied range
    Application.CutCopyMode = False

End Sub'
