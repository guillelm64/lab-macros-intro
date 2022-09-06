Sub getting_nice()
'
' Setting data the desire way
'

'
    Range("A1:E203").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Columns("A:A").ColumnWidth = 14
    Columns("B:B").ColumnWidth = 22.9
    Columns("C:C").ColumnWidth = 34
    Columns("D:D").ColumnWidth = 20.7
    Selection.ColumnWidth = 23.5
    Range("A1:E1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent2
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("E2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.NumberFormat = "mm/dd/yy hh:mm"
End Sub

