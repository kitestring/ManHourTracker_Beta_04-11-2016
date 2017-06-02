VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Methods"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const ActivePW As String = "81643" 'Also in RawData

Sub LockUnlockWb(ByVal WkBk As Workbook, ByVal LockOrUnlock As String, ByVal WkBkType As String)
    Dim WkBkSheets() As String
    Dim i As Byte
    Dim v As Integer
    
    If LockOrUnlock = "lock" Then
        v = 2
    ElseIf LockOrUnlock = "unlock" Then
        v = -1
    End If
    
    If WkBkType = "Control" Then
        Call arrayControlSheet(WkBkSheets)
    ElseIf WkBkType = "DB" Then
        Call arrayDBSheet(WkBkSheets)
    End If
    
    WkBk.Activate
    
    For i = 0 To UBound(WkBkSheets)
    
        If LockOrUnlock = "lock" Then Call LockUnlockWorkbook(WkBkSheets(i), LockOrUnlock)
    
        If i = 2 And WkBkType = "Control" Then
            Call SheetVisablity(WkBkSheets(i), v)
        ElseIf WkBkType = "DB" Then
            Call SheetVisablity(WkBkSheets(i), v)
        End If
        
        If LockOrUnlock = "unlock" Then Call LockUnlockWorkbook(WkBkSheets(i), LockOrUnlock)
        
    Next i
End Sub

Sub SheetVisablity(ByVal SheetName As String, ByVal Visability As Integer)
    '-1 visable
    '2 very hidden
    Sheets(SheetName).Visible = Visability
End Sub

Private Sub arrayDBSheet(ByRef SheetName() As String)
    ReDim SheetName(2) As String
    SheetName(0) = "Template"
    SheetName(1) = "Summary"
    SheetName(2) = "ID_List"
End Sub

Private Sub arrayControlSheet(ByRef SheetName() As String)
    ReDim SheetName(2) As String
    SheetName(0) = "TimeLogger"
    SheetName(1) = "UploadTracker"
    SheetName(2) = "MetaData"
End Sub

Sub LockUnlockWorkbook(ByVal SheetName As String, ByVal LockOrUnlock As String)
    Sheets(SheetName).Select
    If LockOrUnlock = "lock" Then
        ActiveSheet.Protect Password:=ActivePW, DrawingObjects:=True, Contents:=True, Scenarios:=True
    ElseIf LockOrUnlock = "unlock" Then
        ActiveSheet.Unprotect Password:=ActivePW
    End If
End Sub

Sub ReturnRowAndColumn(ByRef row As Integer, ByRef column As Integer, ByVal CellRangeReference As String)
'Assumes you are already in the corresponding workbook
    If Len(CellRangeReference) <> 0 Then
        row = Range(CellRangeReference).row
        column = Range(CellRangeReference).column
    Else
        row = ActiveCell.row
        column = ActiveCell.column
    End If
End Sub

Sub EventLogger(ByVal message As String, ByVal WkBk As Workbook, ByVal SheetName As String)
    'Dim d As String
    'd = FormatDateTime(Date) & " " & FormatDateTime(Date, vbLongTime)
    WkBk.Activate
    Sheets(SheetName).Select
    'Range("EventLog").Value = d & "  ---  " & message & Chr(10) & Range("EventLog").Value
    Range("EventLog").Value = message & Chr(10) & Range("EventLog").Value
    Range("A1").Select
End Sub

Sub OpenFile(ByVal path As String)
    Dim RefWkBk As Workbook
    Workbooks.Open Filename:=path
    Set RefWkBk = ActiveWorkbook
    RefWkBk.Activate
End Sub

Function SheetVisableAndUnlock(ByVal WkBk As Workbook, ByVal SheetName As String) As Boolean
On Error GoTo ErrorCatch
    WkBk.Activate
    Sheets(SheetName).Visible = -1
    Sheets(SheetName).Select
    ActiveSheet.Unprotect Password:=ActivePW
    SheetVisableAndUnlock = True
    Exit Function
ErrorCatch:
    SheetVisableAndUnlock = False
End Function

Sub CloseFile(ByVal WkBk As Workbook)
    WkBk.Close False
End Sub

Sub AutoFitDumpColumns(ByVal c1 As String, ByVal c2 As String)
    Columns(c1 & ":" & c2).EntireColumn.AutoFit
End Sub

Sub AddHeaders(ByVal DumpRow As Integer, ByVal DumpColumn As Integer, ByVal LabelHeader As String, ByVal ValueHeader As String, ByVal PopLabel As String)
    Cells(DumpRow - 2, DumpColumn).Value = PopLabel
    Cells(DumpRow - 1, DumpColumn).Value = LabelHeader
    Cells(DumpRow - 1, DumpColumn + 1).Value = ValueHeader
End Sub

Sub MergePopulationCells(ByVal DumpRow As Integer, ByVal DumpColumn As Integer)
    Range(Cells(DumpRow - 2, DumpColumn), Cells(DumpRow - 2, DumpColumn + 1)).Merge
End Sub

Private Sub BoldHeaderCells(ByVal DumpRow As Integer, ByVal DumpColumn As Integer)
    Range(Cells(DumpRow - 2, DumpColumn), Cells(DumpRow - 1, DumpColumn + 1)).Font.Bold = True
End Sub

Sub SheetWhite()
    Cells.Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub

Private Sub CenterFont()
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
    End With
End Sub


Sub FormatTable_Data(ByVal DumpRow As Integer, ByVal DumpColumn As Integer, ByVal LastRow As String)
    
    Range(Cells(DumpRow, DumpColumn), Cells(LastRow, DumpColumn + 1)).Select
    
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    Call CenterFont
    
End Sub

Sub FormatTable_Headers(ByVal DumpRow As Integer, ByVal DumpColumn As Integer)
    
    Range(Cells(DumpRow - 2, DumpColumn), Cells(DumpRow - 1, DumpColumn + 1)).Select
    
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.14996795556505
        .PatternTintAndShade = 0
    End With
    
    Call CenterFont
    
    Call BoldHeaderCells(DumpRow, DumpColumn)
    
End Sub

Sub ChartTitle(ByVal DumpRow As Integer, ByVal DumpColumn As Integer, ByVal PopulationLabel As String, ByVal CategoryLabel As String)
    Dim Equation As String
    
    Equation = "Hours: " & PopulationLabel & " - " & CategoryLabel
    Cells(DumpRow, DumpColumn + 2).Value = Equation
    Cells(DumpRow, DumpColumn + 2).Select
    
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
End Sub

Sub BuildChart(ByVal DumpRow As Integer, ByVal LastRow As Integer, SubCatSelected As Boolean, ByVal c1 As String, ByVal c2 As String, ByVal c3 As String)
    Dim TopRow As Integer
    Dim SeriesName As String
    Dim SeriesValues As String
    Dim SeriesXValues As String
    
    If SubCatSelected = False Then
        TopRow = DumpRow
    ElseIf SubCatSelected = True Then
        TopRow = DumpRow + 1
    End If
    
    Cells(1, 1).Select
    SeriesName = "=Sheet1!" & c3 & DumpRow
    SeriesValues = "=Sheet1!" & c2 & TopRow & ":" & c2 & LastRow
    SeriesXValues = "=Sheet1!" & c1 & TopRow & ":" & c1 & LastRow
    
    ActiveSheet.Shapes.AddChart2(286, xl3DColumnClustered).Select
    ActiveSheet.Shapes("Chart 1").IncrementLeft -401.25
    ActiveSheet.Shapes("Chart 1").IncrementTop -120.75
    ActiveSheet.Shapes("Chart 1").ScaleWidth 1.93125, msoFalse, msoScaleFromTopLeft
    ActiveSheet.Shapes("Chart 1").ScaleHeight 1.9895833333, msoFalse, _
        msoScaleFromTopLeft
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(1).name = SeriesName
    ActiveChart.FullSeriesCollection(1).Values = SeriesValues
    ActiveChart.FullSeriesCollection(1).XValues = SeriesXValues
End Sub

Sub FormatChart()
    ActiveSheet.ChartObjects("Chart 1").Activate
    ActiveSheet.ChartObjects("Chart 1").Activate
    With ActiveSheet.Shapes("Chart 1").Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
    End With
    With ActiveSheet.Shapes("Chart 1").Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorText1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
    End With
    Application.CommandBars("Format Object").Visible = False
    With ActiveSheet.Shapes("Chart 1").Line
        .Visible = msoTrue
        .Weight = 1.5
    End With
End Sub

Sub AddAxisLabel(ByVal AxisLabel As String)

    ActiveSheet.ChartObjects("Chart 1").Activate
    ActiveChart.Walls.Select
    ActiveChart.SetElement (msoElementPrimaryValueAxisTitleAdjacentToAxis)
    ActiveChart.Axes(xlValue, xlPrimary).AxisTitle.Text = AxisLabel
    Selection.Format.TextFrame2.TextRange.Characters.Text = AxisLabel
    With Selection.Format.TextFrame2.TextRange.Characters(1, 5).ParagraphFormat
        .TextDirection = msoTextDirectionLeftToRight
        .Alignment = msoAlignCenter
    End With
    With Selection.Format.TextFrame2.TextRange.Characters(1, 5).Font
        .BaselineOffset = 0
        .Bold = msoFalse
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(89, 89, 89)
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 10
        .Italic = msoFalse
        .Kerning = 12
        .name = "+mn-lt"
        .UnderlineStyle = msoNoUnderline
        .Strike = msoNoStrike
    End With
   
End Sub

Sub PopulateColumnLetterArray(ByRef ColumnNumber() As String)

    ColumnNumber(1) = "A"
    ColumnNumber(2) = "B"
    ColumnNumber(3) = "C"
    ColumnNumber(4) = "D"
    ColumnNumber(5) = "E"
    ColumnNumber(6) = "F"
    ColumnNumber(7) = "G"
    ColumnNumber(8) = "H"
    ColumnNumber(9) = "I"
    ColumnNumber(10) = "J"
    ColumnNumber(11) = "K"
    ColumnNumber(12) = "L"
    ColumnNumber(13) = "M"
    ColumnNumber(14) = "N"
    ColumnNumber(15) = "O"
    ColumnNumber(16) = "P"
    ColumnNumber(17) = "Q"
    ColumnNumber(18) = "R"
    ColumnNumber(19) = "S"
    ColumnNumber(20) = "T"
    ColumnNumber(21) = "U"
    ColumnNumber(22) = "V"
    ColumnNumber(23) = "W"
    ColumnNumber(24) = "X"
    ColumnNumber(25) = "Y"
    ColumnNumber(26) = "Z"
    
    ColumnNumber(27) = "AA"
    ColumnNumber(28) = "AB"
    ColumnNumber(29) = "AC"
    ColumnNumber(30) = "AD"
    ColumnNumber(31) = "AE"
    ColumnNumber(32) = "AF"
    ColumnNumber(33) = "AG"
    ColumnNumber(34) = "AH"
    ColumnNumber(35) = "AI"
    ColumnNumber(36) = "AJ"
    ColumnNumber(37) = "AK"
    ColumnNumber(38) = "AL"
    ColumnNumber(39) = "AM"
    ColumnNumber(40) = "AN"
    ColumnNumber(41) = "AO"
    ColumnNumber(42) = "AP"
    ColumnNumber(43) = "AQ"
    ColumnNumber(44) = "AR"
    ColumnNumber(45) = "AS"
    ColumnNumber(46) = "AT"
    ColumnNumber(47) = "AU"
    ColumnNumber(48) = "AV"
    ColumnNumber(49) = "AW"
    ColumnNumber(50) = "AX"
    ColumnNumber(51) = "AY"
    ColumnNumber(52) = "AZ"
    
    ColumnNumber(53) = "BA"
    ColumnNumber(54) = "BB"
    ColumnNumber(55) = "BC"
    ColumnNumber(56) = "BD"
    ColumnNumber(57) = "BE"
    ColumnNumber(58) = "BF"
    ColumnNumber(59) = "BG"
    ColumnNumber(60) = "BH"
    ColumnNumber(61) = "BI"
    ColumnNumber(62) = "BJ"
    ColumnNumber(63) = "BK"
    ColumnNumber(64) = "BL"
    ColumnNumber(65) = "BM"
    ColumnNumber(66) = "BN"
    ColumnNumber(67) = "BO"
    ColumnNumber(68) = "BP"
    ColumnNumber(69) = "BQ"
    ColumnNumber(70) = "BR"
    ColumnNumber(71) = "BS"
    ColumnNumber(72) = "BT"
    ColumnNumber(73) = "BU"
    ColumnNumber(74) = "BV"
    ColumnNumber(75) = "BW"
    ColumnNumber(76) = "BX"
    ColumnNumber(77) = "BY"
    ColumnNumber(78) = "BZ"
    
    ColumnNumber(79) = "CA"
    ColumnNumber(80) = "CB"
    ColumnNumber(81) = "CC"
    ColumnNumber(82) = "CD"
    ColumnNumber(83) = "CE"
    ColumnNumber(84) = "CF"
    ColumnNumber(85) = "CG"
    ColumnNumber(86) = "CH"
    ColumnNumber(87) = "CI"
    ColumnNumber(88) = "CJ"
    ColumnNumber(89) = "CK"
    ColumnNumber(90) = "CL"
    ColumnNumber(91) = "CM"
    ColumnNumber(92) = "CN"
    ColumnNumber(93) = "CO"
    ColumnNumber(94) = "CP"
    ColumnNumber(95) = "CQ"
    ColumnNumber(96) = "CR"
    ColumnNumber(97) = "CS"
    ColumnNumber(98) = "CT"
    ColumnNumber(99) = "CU"
    ColumnNumber(100) = "CV"
    ColumnNumber(101) = "CW"
    ColumnNumber(102) = "CX"
    ColumnNumber(103) = "CY"
    ColumnNumber(104) = "CZ"
    
    ColumnNumber(105) = "DA"
    ColumnNumber(106) = "DB"
    ColumnNumber(107) = "DC"
    ColumnNumber(108) = "DD"
    ColumnNumber(109) = "DE"
    ColumnNumber(110) = "DF"
    ColumnNumber(111) = "DG"
    ColumnNumber(112) = "DH"
    ColumnNumber(113) = "DI"
    ColumnNumber(114) = "DJ"
    ColumnNumber(115) = "DK"
    ColumnNumber(116) = "DL"
    ColumnNumber(117) = "DM"
    ColumnNumber(118) = "DN"
    ColumnNumber(119) = "DO"
    ColumnNumber(120) = "DP"
    ColumnNumber(121) = "DQ"
    ColumnNumber(122) = "DR"
    ColumnNumber(123) = "DS"
    ColumnNumber(124) = "DT"
    ColumnNumber(125) = "DU"
    ColumnNumber(126) = "DV"
    ColumnNumber(127) = "DW"
    ColumnNumber(128) = "DX"
    ColumnNumber(129) = "DY"
    ColumnNumber(130) = "DZ"
    
    ColumnNumber(131) = "EA"
    ColumnNumber(132) = "EB"
    ColumnNumber(133) = "EC"
    ColumnNumber(134) = "ED"
    ColumnNumber(135) = "EE"
    ColumnNumber(136) = "EF"
    ColumnNumber(137) = "EG"
    ColumnNumber(138) = "EH"
    ColumnNumber(139) = "EI"
    ColumnNumber(140) = "EJ"
    ColumnNumber(141) = "EK"
    ColumnNumber(142) = "EL"
    ColumnNumber(143) = "EM"
    ColumnNumber(144) = "EN"
    ColumnNumber(145) = "EO"
    ColumnNumber(146) = "EP"
    ColumnNumber(147) = "EQ"
    ColumnNumber(148) = "ER"
    ColumnNumber(149) = "ES"
    ColumnNumber(150) = "ET"
    ColumnNumber(151) = "EU"
    ColumnNumber(152) = "EV"
    ColumnNumber(153) = "EW"
    ColumnNumber(154) = "EX"
    ColumnNumber(155) = "EY"
    ColumnNumber(156) = "EZ"
    
End Sub

