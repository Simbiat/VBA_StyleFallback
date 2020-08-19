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
        'Doing Borders without specification may seem to result in wrong application of the style in some cases, thus doing it for each border separately
        'Also doing them in a loop with some checks for a small optimization
        Dim ArrayElement As Variant
        For Each ArrayElement In Array(xlLeft, xlTop, xlRight, xlBottom, xlDiagonalDown, xlDiagonalUp, xlInsideHorizontal, xlInsideVertical)
            'Line style first
            .Borders(ArrayElement).LineStyle = ThisWorkbook.Styles(StyleName).Borders(ArrayElement).LineStyle
            'Check if it's not "no border", because if it is and we apply color or weight, it will change to Thin
            If .Borders(ArrayElement).LineStyle <> xlLineStyleNone Then
                'Color
                .Borders(ArrayElement).Color = ThisWorkbook.Styles(StyleName).Borders(ArrayElement).Color
                'Weight
                .Borders(ArrayElement).Weight = ThisWorkbook.Styles(StyleName).Borders(ArrayElement).Weight
            End If
        Next ArrayElement
        'Formatting
        .NumberFormat = ThisWorkbook.Styles(StyleName).NumberFormat
        .NumberFormatLocal = ThisWorkbook.Styles(StyleName).NumberFormatLocal
        .FormulaHidden = ThisWorkbook.Styles(StyleName).FormulaHidden
        'Protection
        .Locked = ThisWorkbook.Styles(StyleName).Locked
        On Error GoTo 0
    End With
End Sub
