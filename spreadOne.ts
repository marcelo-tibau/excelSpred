/* Codes in typescript to copy and paste an excel range */
Sub Range_Copy_Examples()

/* Copy from one cell to another */
 Range("A1").Copy Range("C1")

/* Copy from one range to another */
 Range("A1:A3").Copy Range("D1:D3")

/* Copy from a range to a cell */
 Range("A1:A3").Copy Range("D1")

/* Copy from one worksheet to another */
 Worksheets("Sheet1").Range("A1").Copy Worksheets("Sheet2").Range("A1")

/* Copy from one workbook to another */
 Workbooks("Book1.xlsx").Worksheets("Sheet1").Range("A1").Copy _ 
    Workbooks("Book2.xlsx").Worksheets("Sheet1").Range("A1")

End Sub


Sub Paste_Values_Examples()

/* Set the cells values equal to one another */
 Range("C1").Value = Range("A1").Value

 Range("D1:D3").Value = Range("A1:A3").Value

  Workbooks("Book2.xlsx").Worksheets("Sheet1").Range("A1").Value = _ 
        Workbooks("Book1.xlsx").Worksheets("Sheet1").Range("A1").Value


End Sub


Sub PasteSpecial_Examples()
/* Range.PasteSpecial method */
/* Copy and PasteSpecial - between worksheets - between Workbooks - Disable marching ants around copied range */
 Range("A1").Copy
 Range("A3").PasteSpecial

 Worksheets("Sheet1").Range("A2").Copy
 Worksheets("Sheet2").Range("A2").PasteSpecial

 Workbooks("Book1.xlsx").Worksheets("Sheet1").Range("A1").Copy
 Workbooks("Book2.xlsx").Worksheets("Sheet1").Range("A1").PasteSpecial

 Application.CutCopyMode = False

End Sub