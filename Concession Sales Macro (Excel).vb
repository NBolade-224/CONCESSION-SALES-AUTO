
Sub ConcessionSalesMacro()
'
' Macro9 Macro
'

'
    Application.ScreenUpdating = False
    Dim lr As Long
    lr = Cells(Rows.Count, "J").End(xlUp).Row

    Dim CwB As String
    CwB = ActiveWorkbook.Name

    Workbooks.Open "C:\Users\nickb\Videos\Commercial Testing\X by X (SKU, GTIN FOR ZZ) - Copy.xlsx"

    Windows(CwB).Activate


    ActiveWorkbook.Worksheets(1).Sort.SortFields.Add2 Key:= _
        Range("N2:N" & lr), SortOn:=xlSortOnValues, Order:=xlDescending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets(1).Sort
        .SetRange Range("A1:Z" & lr)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Dim i As Integer
    i = 2
    Do While Cells(i, 14).Value <> 1
        Cells(i, 14).Select
        Selection.Delete Shift:=xlToLeft
        i = i + 1
    Loop


    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("B2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-1]<>"""",TEXT(RC[-1],""0000000000000""),"""")"
    Range("B2").Select
    Selection.AutoFill Destination:=Range("B2:B" & lr)
    Range("B2:B" & lr).Select
    Selection.Copy
    Range("A2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("B:B").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft


    ActiveWorkbook.Worksheets(1).Sort.SortFields.Add2 Key:= _
        Range("A2:A" & lr), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets(1).Sort
        .SetRange Range("A1:Z" & lr)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    i = 2
    Do While Cells(i, 1).Value = ""
        Cells(i, 1).FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[14],'[X by X (SKU, GTIN FOR ZZ) - Copy.xlsx]GTIN (NON-PRIMARY)'!C9:C10,2,FALSE),VLOOKUP(RC[14],'[X by X (SKU, GTIN FOR ZZ) - Copy.xlsx]GTIN (PRIMARY UPC)'!C9:C10,2,FALSE))"
        i = i + 1
    Loop


    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("B2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-1]<>"""",TEXT(RC[-1],""0000000000000""),"""")"
    Range("B2").Select
    Selection.AutoFill Destination:=Range("B2:B" & lr)
    Range("B2:B" & lr).Select
    Selection.Copy
    Range("A2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("B:B").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft


    Columns("C:C").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("C2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-1]<>"""",TEXT(RC[-1],""000""),"""")"
    Range("C2").Select
    Selection.AutoFill Destination:=Range("C2:C" & lr)
    Range("C2:C" & lr).Select
    Selection.Copy
    Range("B2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("C:C").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft


    Columns("D:D").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("D2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-1]<>"""",TEXT(RC[-1],""0""),"""")"
    Range("D2").Select
    Selection.AutoFill Destination:=Range("D2:D" & lr)
    Range("D2:D" & lr).Select
    Selection.Copy
    Range("C2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("D:D").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft


    Columns("G:G").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("G2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-1]<>"""",TEXT(RC[-1],""00000000""),"""")"
    Range("G2").Select
    Selection.AutoFill Destination:=Range("G2:G" & lr)
    Range("G2:G" & lr).Select
    Selection.Copy
    Range("F2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("G:G").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft


    Columns("H:H").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("H2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-1]<>"""",TEXT(RC[-1],""000000000""),"""")"
    Range("H2").Select
    Selection.AutoFill Destination:=Range("H2:H" & lr)
    Range("H2:H" & lr).Select
    Selection.Copy
    Range("G2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("H:H").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft


    Columns("L:L").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("L2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-1]<>"""",TEXT(RC[-1],""00""),"""")"
    Range("L2").Select
    Selection.AutoFill Destination:=Range("L2:L" & lr)
    Range("L2:L" & lr).Select
    Selection.Copy
    Range("K2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("L:L").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft


    Columns("M:M").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("M2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-1]<>"""",TEXT(RC[-1],""00""),"""")"
    Range("M2").Select
    Selection.AutoFill Destination:=Range("M2:M" & lr)
    Range("M2:M" & lr).Select
    Selection.Copy
    Range("L2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("M:M").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft


    Columns("Y:Y").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("Y2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-1]<>"""",0,"""")"
    Range("Y2").Select
    Selection.AutoFill Destination:=Range("Y2:Y" & lr)
    Range("Y2:Y" & lr).Select
    Selection.Copy
    Range("X2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("Y:Y").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft


    Columns("E:E").Select
    Selection.NumberFormat = "d-mmm-yy"

    Workbooks("X by X (SKU, GTIN FOR ZZ) - Copy.xlsx").Close SaveChanges:=False
    Application.ScreenUpdating = True

End Sub

