VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmQuery 
   Caption         =   "Time Log Report Builder"
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5370
   OleObjectBlob   =   "frmQuery.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const DumpRow As Integer = 4
Const DumpColumn As Integer = 16 'P
Dim CatAbv() As String

Private Sub UserForm_Initialize()
    'Application.ScreenUpdating = False
    'Application.DisplayAlerts = False
    Dim M As New Methods
    'Call M.LockUnlockWb(ActiveWorkbook, "unlock", "Control")
    
    Dim R As Integer
    Dim c As Integer
    Dim n As Integer
    Dim i As Integer
    
    n = Range("EmployeeCount").Value
    Call M.ReturnRowAndColumn(R, c, "EmployeeName")
    
    Sheets("MetaData").Select
    
    For i = 0 To n
        lstPopulation.AddItem Cells(R + 1 + i, c).Value
    Next i
    
    n = Range("CategoryCount").Value
    Call M.ReturnRowAndColumn(R, c, "Category")
    ReDim CatAbv(n) As String
    
    For i = 0 To n
        lstCategory.AddItem Cells(R + 1 + i, c).Value
        CatAbv(i) = Cells(R + 1 + i, c).Value
    Next i
    
    lstPopulation.Value = "All"
    lstCategory.Value = "All"

    'Call M.LockUnlockWb(ActiveWorkbook, "lock", "Control")
    Sheets("TimeLogger").Select
    'Application.ScreenUpdating = True
    'Application.DisplayAlerts = True
End Sub

Private Sub cmdSetDB_Click()
    Dim M As New Methods
    Dim MinerWkBk As Workbook
    Set MinerWkBk = ActiveWorkbook
    Dim DBfile As Variant
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    DBfile = Application.GetOpenFilename(Title:="Select Employee Time Sheet(s)", MultiSelect:=False, _
        FileFilter:="Excel Workbook (*.xlsx),*.xlsx")
    If VarType(DBfile) = vbBoolean Then
        Call M.EventLogger("Fail  ---  Data base not set.", MinerWkBk, "TimeLogger")
        Exit Sub
    Else
        Range("DB_Directory").Value = DBfile
        Call M.EventLogger("Pass  ---  Data base set to:" & Chr(10) & "     " & DBfile, MinerWkBk, "TimeLogger")
    End If
    
    Sheets("TimeLogger").Select
    Unload frmQuery
End Sub

Private Sub cmdBuildReport_Click()
      
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    frmQuery.Hide
    
    Dim M As New Methods
    
    Dim ColumnLetter(1 To 156) As String
    Call M.PopulateColumnLetterArray(ColumnLetter)
    
    Dim MinerWkBk As Workbook
    Const MineSheet As String = "TimeLogger"
    Const MDSheet As String = "MetaData"
    Const TrackerSheet As String = "UploadTracker"
    Set MinerWkBk = ActiveWorkbook

    Dim DB As New DataBase
    DB.path = Range("DB_Directory").Value
    DB.Summary_Sheet = "Summary"
    
    'Determine if the selected category is All or a sub-category
    'From that information define the source data column
    Dim SubCatSelected As Boolean
    Dim LabelColumn As Integer
    Dim CategoryLabel As String
    
    If lstCategory.ListIndex > 0 Then
        SubCatSelected = True
        LabelColumn = 3
        CategoryLabel = lstCategory.Value
    Else
        SubCatSelected = False
        LabelColumn = 2
        CategoryLabel = "All Categories"
    End If
    
    'From the above info populate the source data row arrays
    Call DB.RowsAndColumns(SubCatSelected, lstCategory.ListIndex - 1)
    
    'Determine if the selected population is All or an individual
    'From that information define the source data sheet
    Dim SourceDataSheet As String
    Dim StringFound As Boolean
    Dim R As Integer
    Dim c As Integer
    Dim PopulationLabel As String
    If lstPopulation.ListIndex > 0 Then
        PopulationLabel = lstPopulation.Value
        Sheets(MDSheet).Select
        StringFound = DB.FindString(lstPopulation.Value)
        Call M.ReturnRowAndColumn(R, c, "")
        SourceDataSheet = Cells(R, c + 2).Value
    Else
        SourceDataSheet = DB.Summary_Sheet
        PopulationLabel = "Entire Team"
    End If
    
    'Check if DB path is valid and open the DB
    Dim DB_file_exists As Boolean
    DB_file_exists = DB.file_exists
    If DB_file_exists = True Then
        'If DB file exists then open it and set ID sheet to visable
        Call M.OpenFile(DB.path)
        DB.WkBk = ActiveWorkbook
        'Call M.LockUnlockWb(DB.WkBk, "unlock", "DB")
    Else
        'If not then inform the user and exit
        MsgBox "The Data Base path is invalid or you don't not" & Chr(10) & _
            "have the correct permissions to access the Data Base." & Chr(10) & _
            "Terminate Time Log Data Miner.", vbCritical, "INVALID DATA BASE PATH"
            Call M.EventLogger("Fail  ---  The Data Base path is invalid or you don't not " & _
                    "have the correct permissions to access the Data Base. Thus, the Time Log Data Miner was terminated.", MinerWkBk, MineSheet)
        Unload frmQuery
        Exit Sub
    End If
    
    'Unlock SourceDataSheet sheet
    'If error is thrown it indicates that there is no data for the selected individual
    Dim SheetExists As Boolean
    SheetExists = M.SheetVisableAndUnlock(DB.WkBk, SourceDataSheet)
    If SheetExists = False Then
        Call M.CloseFile(DB.WkBk)
        Call M.EventLogger("Fail  ---  No data was found within the data base for " & lstPopulation.Value & ".", MinerWkBk, MineSheet)
        Unload frmQuery
        Exit Sub
    End If
    
    'Mine defined data
    Call DB.MineDataBaseSelection(SourceDataSheet, SubCatSelected, LabelColumn)
    Call M.CloseFile(DB.WkBk)

    'Create new workbook
    Dim ReportWkBk As Workbook
    Workbooks.Add
    Set ReportWkBk = ActiveWorkbook
    
    'Dump data into Report workbook
    Dim LastRow As Integer
    If SubCatSelected = True Then
        LastRow = DB.DumpReportData(ReportWkBk, DumpRow, DumpColumn, SubCatSelected, CatAbv)
    ElseIf SubCatSelected = False Then
        LastRow = DB.DumpReportData(ReportWkBk, DumpRow, DumpColumn, SubCatSelected, CatAbv)
    End If
    
    'Label and format Dump Columns
    Call M.SheetWhite
    Call M.AddHeaders(DumpRow, DumpColumn, CategoryLabel, "Hours", PopulationLabel)
    Call M.MergePopulationCells(DumpRow, DumpColumn)
    Call M.AutoFitDumpColumns(ColumnLetter(DumpColumn), ColumnLetter(DumpColumn + 1))
    Call M.FormatTable_Data(DumpRow, DumpColumn, LastRow)
    Call M.FormatTable_Headers(DumpRow, DumpColumn)
    
    'Create and format chart
    Call M.ChartTitle(DumpRow, DumpColumn, PopulationLabel, CategoryLabel)
    Call M.BuildChart(DumpRow, LastRow, SubCatSelected, ColumnLetter(DumpColumn), ColumnLetter(DumpColumn + 1), ColumnLetter(DumpColumn + 2))
    Call M.FormatChart
    
    'Button up
    Cells(1, 1).Select
    Call M.EventLogger("Pass  ---  Report --- " & PopulationLabel & " - " & CategoryLabel & " --- Successfully created.", MinerWkBk, MineSheet)
    
    Unload frmQuery
End Sub


