Attribute VB_Name = "Module1"
Sub LoopInsertFormatHeaders()
' variable for worksheet number
    Dim ws As Worksheet
    
    ' loop thru each worksheet and insert macros to them
    For Each ws In Worksheets
        Worksheets(ws.Name).Select
        
        If Range("A1").Value <> "Division" Then
            InsertHeaders
            FormatHeaders
        End If
    Next ws
End Sub


Sub InsertHeaders()
Attribute InsertHeaders.VB_Description = "Inserts new row and headers of each column"
Attribute InsertHeaders.VB_ProcData.VB_Invoke_Func = " \n14"
'
' InsertHeaders Macro
' Inserts new row and headers of each column
'

'
    Rows("1:1").Select
    Calculate
    Selection.Insert Shift:=xlDown
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Division"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Category"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "Jan"
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "Feb"
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "Mar"
    Range("F1").Select
    ActiveCell.FormulaR1C1 = "Total"
    Range("A1").Select
End Sub


Sub FormatHeaders()
Attribute FormatHeaders.VB_Description = "Formats the header and content"
Attribute FormatHeaders.VB_ProcData.VB_Invoke_Func = " \n14"
'
' FormatHeaders Macro
' Formats the header and content
'

'
    Range("A1:F1").Select
    Selection.Font.Bold = True
    With Selection.Font
        .Name = "Calibri"
        .Size = 15
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    With Selection.Font
        .Name = "Calibri"
        .Size = 12
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    Range("C2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range("C1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("C2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.NumberFormat = _
        "_-[$$-en-CA]* #,##0.00_-;-[$$-en-CA]* #,##0.00_-;_-[$$-en-CA]* ""-""??_-;_-@_-"
    Selection.NumberFormat = _
        "_-[$$-en-CA]* #,##0.0_-;-[$$-en-CA]* #,##0.0_-;_-[$$-en-CA]* ""-""??_-;_-@_-"
    Selection.NumberFormat = _
        "_-[$$-en-CA]* #,##0_-;-[$$-en-CA]* #,##0_-;_-[$$-en-CA]* ""-""??_-;_-@_-"
    Range("B1").Select
    Columns("B:B").EntireColumn.AutoFit
    Range("A2").Select
End Sub
