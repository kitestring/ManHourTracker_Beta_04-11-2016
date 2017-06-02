VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RawData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const RDCurrent As String = "Weekly Time Log"
Dim RDpath As String
Dim RDwkbk As Workbook
Dim RDwkbk2 As Workbook
Dim RDsheet As String
Dim RDmonday As Date
Dim RDverifier As String
Dim RDid As String
Dim RDstartrow As Integer
Dim RDrow(110) As Integer
Const ActivePW As String = "81643" 'Also in Methods

Public Property Let WkBk(RawData_Workbook As Workbook)
    Set RDwkbk = RawData_Workbook
End Property

Public Property Get WkBk() As Workbook
    Set WkBk = RDwkbk
End Property


Public Property Let startrow(RD_startrow As Integer)
    RDstartrow = RD_startrow
End Property

Public Property Get startrow() As Integer
    startrow = RDstartrow
End Property


Public Property Let ID(RD_id As String)
    RDid = RD_id
End Property

Public Property Get ID() As String
    ID = RDid
End Property


Public Property Let verifier(RD_verifier As String)
    RDverifier = RD_verifier
End Property

Public Property Get verifier() As String
    verifier = RDverifier
End Property


Public Property Let path(RD_path As String)
    RDpath = RD_path
End Property

Public Property Get path() As String
    path = RDpath
End Property


Public Property Let Sheet(RD_sheet As String)
    RDsheet = RD_sheet
End Property

Public Property Get Sheet() As String
    Sheet = RDsheet
End Property


Public Property Let monday(RD_monday As Date)
    RDmonday = RD_monday
End Property

Public Property Get monday() As Date
    monday = RDmonday
End Property

Sub SaveAs_Close()
    Application.ScreenUpdating = False
    Dim directory As String
    Dim name As String
    Dim fso As New Scripting.FileSystemObject
    
    directory = RDwkbk.path & "\"
    name = fso.GetBaseName(RDwkbk.name)
    name = name & "_MARKED"
    
    RDwkbk.Activate
    ActiveWorkbook.SaveAs Filename:=directory & name & ".xlsx"
    RDwkbk.Close

End Sub

Function IDvalid(ByVal DBWkBk As Workbook, ByVal IDsheet As String, ByRef message As String) As Boolean
    
    IDvalid = True
    
    'Check if the 1st character is a chr(39)
    'If yes then remove
    If Left(RDid, 1) = Chr(39) Then
        RDid = Right(RDid, Len(RDid) - 1)
    End If
    
    'If the ID length is 3 then add a leading 0
    If Len(RDid) = 3 Then
        RDid = "0" & RDid
    End If
    
    'If the ID length != 4 then ID is not valid
    If Len(RDid) <> 4 Then
        IDvalid = False
        message = "The Employee ID is not the correct number of characters."
    End If

    'If the ID is not matchable to the ID list then ID is not valid
    If IDvalid = True Then
        DBWkBk.Activate
        Sheets(IDsheet).Select
        IDvalid = FindString(RDid)
        If IDvalid = True Then
            message = ""
        ElseIf IDvalid = False Then
            message = "The Employee ID provided does not match any of the records in the data base."
        End If
    End If
    
    If IDvalid = False Then Call HighlightCell("J3")

End Function

Function IsMonday(ByRef message As String) As Boolean
    Dim intWeekDay As Integer
    intWeekDay = WeekDay(RDmonday)
    If intWeekDay = 2 Then
        IsMonday = True
    Else
        IsMonday = False
        message = "The date entered in the " & Chr(34) & "Mondays Date" & Chr(34) & " cell is not a Monday."
        Call HighlightCell("J2")
    End If
End Function

Function sheet_verifier_sets_RDstartrow(ByRef message As String, ByVal Control As Workbook) As Boolean
    On Error GoTo ErrorCatch
    RDwkbk.Activate
    RDverifier = Range("Verifier").Value
    Dim v As String
    
    If RDverifier = "!^$dfuj4862" Or RDverifier = "#*26554MHJ~" Then
        
        Dim M As New Methods
        
        If RDverifier = "!^$dfuj4862" Then
            v = "1.0"
        ElseIf RDverifier = "#*26554MHJ~" Then
            sheet_verifier_sets_RDstartrow = True
            v = "1.1"
        End If
        
        sheet_verifier_sets_RDstartrow = True
        RDstartrow = 9
        message = "Data transferred to current version."
        Call M.EventLogger("Legacy template loaded  ---  " & RDwkbk.name & " --- " & message, Control, "TimeLogger")
        Call Transfer2Current(v, Control)
                    
    ElseIf RDverifier = "v1.2" Or RDverifier = "v1.3" Then
        sheet_verifier_sets_RDstartrow = True
        RDstartrow = 9
    Else
        GoTo ErrorCatch
    End If
    
    Exit Function
ErrorCatch:
    message = "Is an invalid file, thus was excluded."
    sheet_verifier_sets_RDstartrow = False
End Function

Private Sub HighlightCell(ByVal ErrorRange As String)
    RDwkbk.Activate
    ActiveSheet.Unprotect (ActivePW)
    Range(ErrorRange).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .Color = -16776961
        .TintAndShade = 0
    End With
    ActiveSheet.Protect Password:=ActivePW, DrawingObjects:=True, Contents:=True, Scenarios:=True
End Sub

Private Function FindString(ByVal str As String) As Boolean
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

Sub MineData(ByRef arrayRD() As String)
    Dim day As Integer
    Dim field As Integer
    Dim column As Integer
    Const starting_column As Integer = 4
    RDwkbk.Activate
    
    For day = 0 To 6
        column = starting_column + day
        For field = 0 To 110
            arrayRD(day, field) = Cells(RDrow(field), column).Value
        Next field
    Next day
    
    End Sub

Private Sub Transfer2Current(ByVal v As String, ByVal WkBk As Workbook)
    Dim M As New Methods
    WkBk.Activate
    Call M.SheetVisablity(RDCurrent, -1)
    Sheets(RDCurrent).Select
    Sheets(RDCurrent).Copy
    Set RDwkbk2 = ActiveWorkbook
    
    Dim rng As String
    Dim Drow(12) As Integer
    Dim Rows(1, 12) As Integer
    Dim i As Integer
    Dim j As Integer
    
    Rows(0, 0) = 7
    Rows(1, 0) = 13
        
    Rows(0, 1) = 15
    Rows(1, 1) = 21
        
    Rows(0, 2) = 23
    Rows(1, 2) = 29
    
    Rows(0, 3) = 31
    Rows(1, 3) = 36
    
    Rows(0, 4) = 38
    Rows(1, 4) = 43
    
    Rows(0, 5) = 45
    Rows(1, 5) = 48
    
    Rows(0, 6) = 50
    Rows(1, 6) = 53
    
    Rows(0, 7) = 55
    Rows(1, 7) = 62
    
    Rows(0, 8) = 64
    Rows(1, 8) = 66
    
    Rows(0, 9) = 68
    Rows(1, 9) = 76
    
    Rows(0, 10) = 78
    Rows(1, 10) = 90
    
    Rows(0, 11) = 92
    Rows(1, 11) = 105
    
    Rows(0, 12) = 107
    Rows(1, 12) = 115
    
    Drow(0) = 10
    Drow(1) = 19
    Drow(2) = 28
    Drow(3) = 37
    Drow(4) = 45
    Drow(5) = 53
    Drow(6) = 59
    Drow(7) = 65
    Drow(8) = 75
    Drow(9) = 80
    Drow(10) = 91
    Drow(11) = 106
    Drow(12) = 122
    
    RDwkbk.Activate
    Range("J2:J3").Copy
    
    RDwkbk2.Activate
    Range("J2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    For i = 0 To 12
        For j = 0 To 1
            If v = "1.1" Then Rows(j, i) = Rows(j, i) + 1
        Next j
        rng = "D" & Rows(0, i) & ":" & "J" & Rows(1, i)
        
        RDwkbk.Activate
        Range(rng).Copy
        
        RDwkbk2.Activate
        Range("D" & Drow(i)).Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Application.CutCopyMode = False
    Next i
    
    Dim directory As String
    Dim name As String
    Dim fso As New Scripting.FileSystemObject
    
    directory = RDwkbk.path & "\"
    name = fso.GetBaseName(RDwkbk.name)
    name = name & "_MARKED"
    
    RDwkbk2.Activate
    ActiveWorkbook.SaveAs Filename:=directory & name & ".xlsx"
    RDwkbk.Close
    
    Set RDwkbk = ActiveWorkbook
    
    WkBk.Activate
    Call M.SheetVisablity(RDCurrent, 2)
    
    RDwkbk.Activate
End Sub

Sub PopulateRowReferences()
    Dim i As Integer
    Dim sr As Integer
    Dim er As Integer
    Dim icount As Integer
    Dim cats As Integer
    Dim c As Integer
    Dim top_offset As Integer
    Dim bot_Offset As Integer
    
    icount = 1
    cats = 12
    
    RDrow(0) = 6
    
    For c = 0 To cats
        Select Case c
            Case 0
                'Installs
                top_offset = 0
                bot_Offset = 7
            Case 1
                'Preventative Maintenance Site Visits
                top_offset = 9
                bot_Offset = 16
            Case 2
                'Instrument Repair or Instrument Troubleshooting at a Customer Site
                top_offset = 18
                bot_Offset = 25
            Case 3
                'Remote Hardware Support
                top_offset = 27
                bot_Offset = 33
            Case 4
                'Remote Software Support
                top_offset = 35
                bot_Offset = 41
            Case 5
                'Hardware Repair, Upgrade, or Refurbish (In-House)
                top_offset = 43
                bot_Offset = 47
            Case 6
                'Miscellaneous
                top_offset = 49
                bot_Offset = 53
            Case 7
                'Document Generation
                top_offset = 55
                bot_Offset = 63
            Case 8
                'Software or Hardware Support Internal (R&D)
                top_offset = 65
                bot_Offset = 68
            Case 9
                'Online Training
                top_offset = 70
                bot_Offset = 79
            Case 10
                'Onsite Training
                top_offset = 81
                bot_Offset = 94
            Case 11
                'In-house Training
                top_offset = 96
                bot_Offset = 110
            Case 12
                'Validation Duties
                top_offset = 112
                bot_Offset = 121
        End Select
        
        sr = top_offset + RDstartrow
        er = bot_Offset + RDstartrow
        
        For i = sr To er
            RDrow(icount) = i
            icount = icount + 1
        Next i
    Next c

End Sub
