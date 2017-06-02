VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim DBpath As String
Dim DBwk As Workbook
Dim DBTemplate As String
Dim DBID As String
Dim DBSummary As String
Dim DBRow(110) As Integer
Dim DBColumn(110) As Integer
Const StaringRow As Integer = 2 'Instal tot row - 2
Const StartingColumn As Integer = 6 'Date column
Const DumpRowOffset As Byte = 4 'Value to add to "=COUNTA(F:F)+4" to find bottom row number
Const NumberOfCategories As Integer = 12
Dim MainCatRow(NumberOfCategories) As Integer
Dim CategoryCount(NumberOfCategories) As Integer
Dim SubCatRow() As Integer
Dim QueryData() As String

'Begin Defining Properties

Public Property Get CatRow() As Integer
    CatRow = MainCatRow
End Property

Public Property Let row(DataBase_Row As Integer)
    DBRow = DataBase_Row
End Property

Public Property Get row() As Integer
    row = DBRow
End Property


Public Property Let column(DataBase_Column As Integer)
    DBColumn = DataBase_Column
End Property

Public Property Get column() As Integer
    column = DBColumn
End Property


Public Property Let WkBk(DataBase_Workbook As Workbook)
    Set DBwk = DataBase_Workbook
End Property

Public Property Get WkBk() As Workbook
    Set WkBk = DBwk
End Property


Public Property Let path(DB_path As String)
    DBpath = DB_path
End Property

Public Property Get path() As String
    path = DBpath
End Property


Public Property Let Template_Sheet(Template As String)
    DBTemplate = Template
End Property

Public Property Get Template_Sheet() As String
    Template_Sheet = DBTemplate
End Property


Public Property Let ID_Sheet(ID As String)
    DBID = ID
End Property

Public Property Get ID_Sheet() As String
    ID_Sheet = DBID
End Property


Public Property Let Summary_Sheet(Summary As String)
    DBSummary = Summary
End Property

Public Property Get Summary_Sheet() As String
    Summary_Sheet = DBSummary
End Property
'End Defining Properties


'Begin Methods

Function file_exists() As Boolean
Dim strResult As String
    strResult = Dir(DBpath)
    If strResult = "" Then
        file_exists = False
    Else
        file_exists = True
    End If
End Function

Sub SheetVeryHidden(ByVal SheetName As String)
    DBwk.Activate
    Sheets(SheetName).Visible = 2
End Sub

Function SheetVisable(ByVal SheetName As String) As Boolean
On Error GoTo ErrorCatch
    DBwk.Activate
    Sheets(SheetName).Visible = -1
    SheetVisable = True
Exit Function
ErrorCatch:
    SheetVisable = False
End Function
Function ID_exists(ByVal ID As String) As String
On Error GoTo ErrorCatch
    If SheetVisable(ID) = False Then GoTo ErrorCatch
    Sheets(ID).Select
    ID_exists = "True"
Exit Function
ErrorCatch:
    Call New_ID(ID)
    ID_exists = "False"
End Function

Private Sub New_ID(ByVal ID As String)
    DBwk.Activate
    Sheets(DBTemplate).Copy After:=DBwk.Sheets(Sheets.Count)
    Sheets(DBTemplate & " (2)").name = ID
    Call AddSummaryLinks(ID)
End Sub

Private Sub AddSummaryLinks(ByVal ID As String)
    Dim i As Integer
    Dim CurrentEquation As String
    Dim EmptyCell As Boolean
    Dim Equation As String
    Dim AppendEquation As String
    
    DBwk.Activate
    Sheets(DBSummary).Select
    CurrentEquation = Range("D4").Formula
    If Len(CurrentEquation) = 0 Then EmptyCell = True
    
    For i = 1 To UBound(DBRow)
        AppendEquation = Chr(39) & ID & Chr(39) & "!D" & DBRow(i) & ")"
        If EmptyCell = True Then
            Equation = "=SUM(" & AppendEquation
        ElseIf EmptyCell = False Then
            CurrentEquation = Cells(DBRow(i), 4).Formula
            Equation = Left(CurrentEquation, Len(CurrentEquation) - 1) & "," & AppendEquation
        End If
            
        Cells(DBRow(i), 4).Value = Equation
        
    Next i
   
    Sheets(ID).Select
    
End Sub

Function DuplicateDates(ByVal monday As String, ByVal ID As String) As String
    Dim boo2str As Boolean
    DBwk.Activate
    Sheets(ID).Select
    boo2str = FindString(monday)
    If boo2str = True Then
        DuplicateDates = "True"
    ElseIf boo2str = False Then
        DuplicateDates = "False"
    End If
    
End Function

Sub ClearDuplateDataSet(ByVal monday As String, ByVal ID As String)
    Dim b As Boolean
    Dim R As Double
    DBwk.Activate
    Sheets(ID).Select
    b = FindString(monday)
    R = ActiveCell.row
    Range(Cells(R, DBColumn(0)), Cells(R + 6, DBColumn(UBound(DBColumn)))).Select
    Selection.Delete Shift:=xlUp
End Sub

Function FindString(ByVal str As String) As Boolean
On Error GoTo ErrorCatch
    
    Cells(1, 1).Select
    Cells.Find(What:=str, After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
    
    FindString = True
    Exit Function
ErrorCatch:
    FindString = False
End Function

Sub CreateTempLinks(ByRef arrayColumnLetter() As String)
    Dim i As Integer
    Dim c As String
    Dim f As String
    
    f = "=SUM(" & c & ":" & c & ")"
    For i = 0 To 109
        c = arrayColumnLetter(DBColumn(i))
        f = "=SUM(" & c & ":" & c & ")"
        Cells(DBRow(i), 4).Value = f
    Next i
End Sub

Sub DumpData(ByRef RDset() As String, ByVal ID As String)
    Dim day As Byte
    Dim field As Byte
    Dim BeginingRow As Double
    Dim row As Double

    BeginingRow = BottomRow(ID)
    For day = 0 To 6
        row = BeginingRow + day
        For field = 0 To 110
            If RDset(day, field) = "" Then RDset(day, field) = "0"
            Cells(row, DBColumn(field)).Value = RDset(day, field)
        Next field
    Next day
    
    Call DataSortByDate(ID)
End Sub

Private Function BottomRow(ByVal SheetName As String) As Double
    DBwk.Activate
    Sheets(SheetName).Select
    Range("A1").Value = "=COUNTA(F:F)+" & DumpRowOffset
    BottomRow = Range("A1").Value
    Range("A1").ClearContents
End Function

Private Sub DataSortByDate(ByVal SheetName As String)
    Dim SortRange As String
    Dim WholeRange As String
    Dim LastRow As Double
    LastRow = BottomRow(SheetName)
    SortRange = "F5:F" & LastRow
    WholeRange = "F5:DY" & LastRow
    
    ActiveWorkbook.Worksheets(SheetName).Sort.SortFields.Clear
    ActiveWorkbook.Worksheets(SheetName).Sort.SortFields.Add Key:=Range(SortRange), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets(SheetName).Sort
        .SetRange Range(WholeRange)
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Range("A1").Select
End Sub

Private Sub DefineCategoryCount()
    CategoryCount(0) = 8 'Installs
    CategoryCount(1) = 8 'Preventative Maintenance Site Visits
    CategoryCount(2) = 8 'Instrument Repair or Instrument Troubleshooting at a Customer Site
    CategoryCount(3) = 7 'Remote Hardware Support
    CategoryCount(4) = 7 'Remote Software Support
    CategoryCount(5) = 5 'Hardware Repair, Upgrade, or Refurbish (In-House)
    CategoryCount(6) = 5 'Miscellaneous
    CategoryCount(7) = 9 'Document Generation
    CategoryCount(8) = 4 'Software or Hardware Support Internal (R&D)
    CategoryCount(9) = 10 'Online Training
    CategoryCount(10) = 14 'Onsite Training
    CategoryCount(11) = 15 'In-house Training
    CategoryCount(12) = 10 'Validation Duties
End Sub

Sub RowsAndColumns(ByVal query As Boolean, Optional ByVal catIndex As Integer)

    Dim cat As Integer
    Dim field As Integer
    Dim i As Integer
    Dim R As Integer
    Dim c As Integer
    Dim j As Integer

    Call DefineCategoryCount
        
    DBColumn(0) = StartingColumn
    i = 1
    j = 0
    
    For cat = 0 To UBound(CategoryCount)

        For field = 0 To CategoryCount(cat) - 1
            R = StaringRow + cat + i + 1
            c = StartingColumn + cat + i + 1
            DBRow(i) = R
            DBColumn(i) = c
            If field = 0 Then
                MainCatRow(cat) = R
            End If
            If query = True And catIndex = cat Then
                If j = 0 Then ReDim SubCatRow(CategoryCount(cat) - 1) As Integer
                SubCatRow(j) = R
                j = j + 1
            End If
            i = i + 1
        Next field
        
    Next cat

End Sub

Sub MineDataBaseSelection(ByVal SheetName As String, ByVal SubCatSelected As Boolean, ByVal LblColumn As Integer)
    Dim i As Integer
    Dim n As Integer
    Dim R As Integer
    
    DBwk.Activate
    Sheets(SheetName).Select
    
    'QueryData(x, y)
        'x = 0 (x-axis label)
        'y = 1 (y-axis value)
    If SubCatSelected = False Then
        n = UBound(MainCatRow)
    ElseIf SubCatSelected = True Then
        n = UBound(SubCatRow)
    End If
    
    ReDim QueryData(1, n) As String
    
    For i = 0 To n
    
        If SubCatSelected = False Then
            R = MainCatRow(i)
        ElseIf SubCatSelected = True Then
            R = SubCatRow(i)
        End If

        QueryData(0, i) = Cells(R, LblColumn).Value
        QueryData(1, i) = Cells(R, 4).Value
        
    Next i

End Sub

Function DumpReportData(ByVal RptWkBk As Workbook, ByVal DumpRow As Integer, ByVal DumpColumn As Integer, ByVal SubCatSelected As Boolean, ByRef CatAbv() As String) As Integer
    RptWkBk.Activate
    Dim i As Integer
    Dim j As Integer
    Dim n As Integer
    Dim column As Integer
    Dim row As Integer
    
    If SubCatSelected = False Then
        n = UBound(MainCatRow)
    ElseIf SubCatSelected = True Then
        n = UBound(SubCatRow)
    End If
    
    For i = 0 To 1
        column = i + DumpColumn
        For j = 0 To n
            row = j + DumpRow
            
            If SubCatSelected = False And i = 0 Then
                Cells(row, column).Value = CatAbv(j + 1)
            Else
                Cells(row, column).Value = QueryData(i, j)
            End If
            
        Next j
    Next i
    DumpReportData = row
End Function
