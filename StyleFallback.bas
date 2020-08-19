Private Sub StyleFallback(ByVal CellRange As Range, StyleName As String)
    If ThisWorkbook.MultiUserEditing = False Then
        'Use common function to apply the style
        'Tracking errors to allow fallback in case of failure
        On Error Resume Next
        CellRange.Style = StyleName
        If Err.Number = 0 Then
            On Error GoTo 0
            Exit Sub
        End If
        On Error GoTo 0
    End If
    'In case we have a shared workbook do this, to simulate the behavior
    With CellRange
        'Have to skipp errors, because a style may be missing certain type of setting in it, which is by design
        On Error Resume Next
        'Content placement
        .IndentLevel = ThisWorkbook.Styles(StyleName).IndentLevel
        .Orientation = ThisWorkbook.Styles(StyleName).Orientation
        .ReadingOrder = ThisWorkbook.Styles(StyleName).ReadingOrder
        .ShrinkToFit = ThisWorkbook.Styles(StyleName).ShrinkToFit
        .HorizontalAlignment = ThisWorkbook.Styles(StyleName).HorizontalAlignment
        .VerticalAlignment = ThisWorkbook.Styles(StyleName).VerticalAlignment
        .WrapText = ThisWorkbook.Styles(StyleName).WrapText
        'Font
        With .Font
            .Bold = ThisWorkbook.Styles(StyleName).Font.Bold
            .Color = ThisWorkbook.Styles(StyleName).Font.Color
            .ColorIndex = ThisWorkbook.Styles(StyleName).Font.ColorIndex
            .FontStyle = ThisWorkbook.Styles(StyleName).Font.FontStyle
            .Italic = ThisWorkbook.Styles(StyleName).Font.Italic
            .Size = ThisWorkbook.Styles(StyleName).Font.Size
            .Strikethrough = ThisWorkbook.Styles(StyleName).Font.Strikethrough
            .Subscript = ThisWorkbook.Styles(StyleName).Font.Subscript
            .Superscript = ThisWorkbook.Styles(StyleName).Font.Superscript
            .ThemeColor = ThisWorkbook.Styles(StyleName).Font.ThemeColor
            .ThemeFont = ThisWorkbook.Styles(StyleName).Font.ThemeFont
            .TintAndShade = ThisWorkbook.Styles(StyleName).Font.TintAndShade
            .Underline = ThisWorkbook.Styles(StyleName).Font.Underline
        End With
        'Background style
        .Interior.Color = ThisWorkbook.Styles(StyleName).Interior.Color
        .Interior.Pattern = ThisWorkbook.Styles(StyleName).Interior.Pattern
        'Doing Borders without specification may seem to result in wrong appliction of the style in some cases, thus doing it for each border separately
        'Color
        .Borders(xlLeft).Color = ThisWorkbook.Styles(StyleName).Borders(xlLeft).Color
        .Borders(xlTop).Color = ThisWorkbook.Styles(StyleName).Borders(xlTop).Color
        .Borders(xlRight).Color = ThisWorkbook.Styles(StyleName).Borders(xlRight).Color
        .Borders(xlBottom).Color = ThisWorkbook.Styles(StyleName).Borders(xlBottom).Color
        .Borders(xlDiagonalDown).Color = ThisWorkbook.Styles(StyleName).Borders(xlDiagonalDown).Color
        .Borders(xlDiagonalUp).Color = ThisWorkbook.Styles(StyleName).Borders(xlDiagonalUp).Color
        .Borders(xlInsideHorizontal).Color = ThisWorkbook.Styles(StyleName).Borders(xlInsideHorizontal).Color
        .Borders(xlInsideVertical).Color = ThisWorkbook.Styles(StyleName).Borders(xlInsideVertical).Color
        'Weight
        .Borders(xlLeft).Weight = ThisWorkbook.Styles(StyleName).Borders(xlLeft).Weight
        .Borders(xlTop).Weight = ThisWorkbook.Styles(StyleName).Borders(xlTop).Weight
        .Borders(xlRight).Weight = ThisWorkbook.Styles(StyleName).Borders(xlRight).Weight
        .Borders(xlBottom).Weight = ThisWorkbook.Styles(StyleName).Borders(xlBottom).Weight
        .Borders(xlDiagonalDown).Weight = ThisWorkbook.Styles(StyleName).Borders(xlDiagonalDown).Weight
        .Borders(xlDiagonalUp).Weight = ThisWorkbook.Styles(StyleName).Borders(xlDiagonalUp).Weight
        .Borders(xlInsideHorizontal).Weight = ThisWorkbook.Styles(StyleName).Borders(xlInsideHorizontal).Weight
        .Borders(xlInsideVertical).Weight = ThisWorkbook.Styles(StyleName).Borders(xlInsideVertical).Weight
        'LineStyle
        'Needs to be applied last or it may add borders, where there are none
        .Borders(xlLeft).LineStyle = ThisWorkbook.Styles(StyleName).Borders(xlLeft).LineStyle
        .Borders(xlTop).LineStyle = ThisWorkbook.Styles(StyleName).Borders(xlTop).LineStyle
        .Borders(xlRight).LineStyle = ThisWorkbook.Styles(StyleName).Borders(xlRight).LineStyle
        .Borders(xlBottom).LineStyle = ThisWorkbook.Styles(StyleName).Borders(xlBottom).LineStyle
        .Borders(xlDiagonalDown).LineStyle = ThisWorkbook.Styles(StyleName).Borders(xlDiagonalDown).LineStyle
        .Borders(xlDiagonalUp).LineStyle = ThisWorkbook.Styles(StyleName).Borders(xlDiagonalUp).LineStyle
        .Borders(xlInsideHorizontal).LineStyle = ThisWorkbook.Styles(StyleName).Borders(xlInsideHorizontal).LineStyle
        .Borders(xlInsideVertical).LineStyle = ThisWorkbook.Styles(StyleName).Borders(xlInsideVertical).LineStyle
        'Formatting
        .NumberFormat = ThisWorkbook.Styles(StyleName).NumberFormat
        .NumberFormatLocal = ThisWorkbook.Styles(StyleName).NumberFormatLocal
        .FormulaHidden = ThisWorkbook.Styles(StyleName).FormulaHidden
        'Protection
        .Locked = ThisWorkbook.Styles(StyleName).Locked
        On Error GoTo 0
    End With
End Sub
