Attribute VB_Name = "TimeLogger"
Option Explicit

Sub Miner()
    
    On Error GoTo ErrorCatch
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Dim MinerWkBk As Workbook
    Const MineSheet As String = "TimeLogger"
    Set MinerWkBk = ActiveWorkbook
    
    'Create a Methods object and unlock controls
    Dim M As New Methods
    Dim ColumnLetter(1 To 156) As String
    Call M.LockUnlockWb(MinerWkBk, "unlock", "Control")
    Call M.PopulateColumnLetterArray(ColumnLetter)
    
    'Create a data base object and construct/initialize it
    Dim DB As New DataBase
    DB.ID_Sheet = "ID_List"
    DB.Summary_Sheet = "Summary"
    DB.Template_Sheet = "Template"
    DB.path = Range("DB_Directory").Value
    Call DB.RowsAndColumns(False)
    
    Dim DB_file_exists As Boolean
    DB_file_exists = DB.file_exists
    If DB_file_exists = True Then
        'If DB file exists then open it and set ID sheet to visable
        Call M.OpenFile(DB.path)
        DB.WkBk = ActiveWorkbook
        Call M.LockUnlockWb(DB.WkBk, "unlock", "DB")
    Else
        'If not then inform the user and exit
        MsgBox "The Data Base path is invalid or you don't not" & Chr(10) & _
            "have the correct permissions to access the Data Base." & Chr(10) & _
            "Terminate Time Log Data Miner.", vbCritical, "INVALID DATA BASE PATH"
            Call M.EventLogger("Fail  ---  The Data Base path is invalid or you don't not " & _
                    "have the correct permissions to access the Data Base. Thus, the Time Log Data Miner was terminated.", MinerWkBk, MineSheet)
        Exit Sub
    End If
    
    'Prompt user for csv files to be data mined
    Dim RawFiles As Variant
    Dim RawFileCount As Integer
    Dim file As Integer
    If MultiFileOpen(RawFiles, RawFileCount) = False Then
        DB.WkBk.Close
        Call M.LockUnlockWb(MinerWkBk, "lock", "Control")
        Sheets(MineSheet).Select
        Exit Sub
    End If
    
    'Create a data base object and construct/initialize it
    Dim RD As New RawData
    RD.Sheet = "Weekly Time Log"
    
    Dim valid_date As Boolean
    Dim valid_id As Boolean
    Dim valid_workbook As Boolean
    Dim error_message As String
    Dim IDsheet_exists As String
    Dim DateFound As String
    Dim RDinfo(6, 110) As String
    Dim yes_no As Byte
    Dim OverwriteData As String
    Dim RDWkBkName As String
    
    For file = 1 To RawFileCount
        
        'Open ith raw data file
        RD.path = RawFiles(file)
        Call M.OpenFile(RD.path)
        RD.WkBk = ActiveWorkbook
        RDWkBkName = RD.WkBk.name
        RD.monday = Range("J2").Value
        RD.ID = Range("J3").Value
        
        'Confirm that the workbook is valid
        valid_workbook = RD.sheet_verifier_sets_RDstartrow(error_message, MinerWkBk)
        If valid_workbook = True Then
            
            'Validate Date
            valid_date = RD.IsMonday(error_message)
            If valid_date = False Then
                Call M.EventLogger("Fail  ---  " & RD.WkBk.name & "  ---  " & error_message & _
                    " The file has been marked and was not imported into the data base. Please make the necessary corrections and reload file.", MinerWkBk, MineSheet)
            End If
            
            'Validate ID
            valid_id = RD.IDvalid(DB.WkBk, DB.ID_Sheet, error_message)
            If valid_id = False Then
                Call M.EventLogger("Fail  ---  " & RD.WkBk.name & "  ---  " & error_message & _
                    " The file has been marked and was not imported into the data base. Please make the necessary corrections and reload file.", MinerWkBk, MineSheet)
            End If
            
            If valid_date = False Or valid_id = False Then RD.SaveAs_Close

        Else
            'If invalid workbook log the even and close
            Call M.EventLogger("Fail  ---  " & RD.WkBk.name & "  ---  " & error_message, MinerWkBk, MineSheet)
            Call M.CloseFile(RD.WkBk)
        End If
        
        'If the workbook and data is valid then mine the data and then dump it into the data base
        If valid_date = True And valid_workbook = True And valid_id = True Then
            RD.PopulateRowReferences 'RDstartrow must be set for this to work
            
            'Mine the data from the raw data work book and then close it
            Call RD.MineData(RDinfo)
            Call M.CloseFile(RD.WkBk)
            
            'Clear logic values to null
            IDsheet_exists = ""
            DateFound = ""
            OverwriteData = ""
            
            'Check if Sheet for given ID exists
            'If not, it will be created and links to summary sheet generated
            IDsheet_exists = DB.ID_exists(RD.ID)
            
            If IDsheet_exists = "True" Then
            
                'Check if ID sheet already has data from the same date range as the mined data
                DateFound = DB.DuplicateDates(RDinfo(0, 0), RD.ID) 'RDinfo(0, 0) Contains mondays data as a string
                
                If DateFound = "True" Then
                
                    'Prompt user to overwrite
                    Call M.EventLogger("Overwrite Warning  ---  " & RDWkBkName & "  ---  " & _
                        "Data was found in the data base for this employee matching the date range in the file loaded.", MinerWkBk, MineSheet)
                    yes_no = MsgBox("Data was found in the data base for this employee" & Chr(10) & _
                        "matching the date range in the file loaded." & Chr(10) & _
                        "Do you wish to overwrite it?", vbYesNo, "OVERWRITE WARNING") 'yes = 6 no = 7
                    
                    If yes_no = 6 Then 'yes = 6
                        'Clear the existing set of data with same dates
                        Call DB.ClearDuplateDataSet(RDinfo(0, 0), RD.ID)
                        OverwriteData = "True"
                        Call M.EventLogger("Data Overwritten  ---  " & RDWkBkName & "  ---  " & _
                            "User overwrite data set.", MinerWkBk, MineSheet)
                    ElseIf yes_no = 7 Then 'no = 7
                        OverwriteData = "False"
                        Call M.EventLogger("Data Overwrite Skipped  ---  " & RDWkBkName & "  ---  " & _
                            "User abort data set overwrite.", MinerWkBk, MineSheet)
                    End If
                    
                End If
            End If
            
            If IDsheet_exists = "False" Or DateFound = "False" Or OverwriteData = "True" Then
                Call DB.DumpData(RDinfo, RD.ID)
                Call M.EventLogger("Pass  ---  " & RDWkBkName & "  ---  File successfully imported into data base.", MinerWkBk, MineSheet)
            End If
            
            Call DB.SheetVeryHidden(RD.ID)
        End If
    Next file
    
    Call M.LockUnlockWb(DB.WkBk, "lock", "DB")
    Call M.LockUnlockWb(MinerWkBk, "lock", "Control")
    Sheets(MineSheet).Select
    Call CloseSaveFile(DB.WkBk)

Exit Sub
ErrorCatch:
    Call M.EventLogger("Unhandled Exception  ---  Please contact software support for assistance.", MinerWkBk, MineSheet)
    Call M.LockUnlockWb(MinerWkBk, "lock", "Control")
    Call M.CloseFile(DB.WkBk)
    MsgBox "An unhandled exception has occurred." & Chr(10) & "Please contact software support for assistance.", vbCritical, "UNHANDLED EXCEPTION"
    
End Sub

Private Function MultiFileOpen(ByRef funRawFile As Variant, ByRef funRawFileCount As Integer) As Boolean
    funRawFile = Application.GetOpenFilename(Title:="Select Employee Time Sheet(s)", MultiSelect:=True, _
        FileFilter:="Excel Workbook (*.xls*),*.xls*")
    If VarType(funRawFile) = vbBoolean Then
        MultiFileOpen = False
        Exit Function
    End If
    funRawFileCount = UBound(funRawFile)
    MultiFileOpen = True
    
End Function

Private Sub CloseSaveFile(ByVal WkBk As Workbook)
    WkBk.Close True
End Sub

