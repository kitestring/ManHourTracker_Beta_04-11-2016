VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UpdateInternational"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const ActivePW As String = "81643"

Sub UpdateDB_International()
'This updates the database headers so the the other sub categories in
'remote software & remote hardware support are changed to internationa.

    Dim M As New Methods
    
    Dim MinerWkBk As Workbook
    Set MinerWkBk = ActiveWorkbook
    MinerWkBk.Activate
    
    Dim DB As New DataBase
    DB.path = Range("DB_Directory").Value
    DB.Summary_Sheet = "Summary"

    Call M.OpenFile(DB.path)
    DB.WkBk = ActiveWorkbook
    Call M.LockUnlockWb(DB.WkBk, "unlock", "DB")
    
    Dim i As Integer
    Dim n As Integer
    Dim rng(1) As String
        rng(0) = "C37"
        rng(1) = "C45"
    Dim SubCatLbl(1) As String
        SubCatLbl(0) = "Intl. Hardware Support"
        SubCatLbl(1) = "Intl. Software Support"
        
    For i = 1 To Sheets.Count
        If Sheets(i).name = "Template" Or Sheets(i).name = "Summary" Then
            Sheets(i).Select
            Range(rng(0)).Select
            Range(rng(0)).Value = SubCatLbl(0)
            Range(rng(1)).Value = SubCatLbl(1)
            Cells(1, 1).Select
        ElseIf Sheets(i).name = "ID_List" Or Sheets(i).name = "DataBase" Then
            n = 1
        Else
            Sheets(i).Visible = -1
            Sheets(i).Select
            ActiveSheet.Unprotect Password:=ActivePW
            
            Range(rng(0)).Select
            Range(rng(0)).Value = SubCatLbl(0)
            Range(rng(1)).Value = SubCatLbl(1)
            Cells(1, 1).Select
            
            ActiveSheet.Protect Password:=ActivePW
            Sheets(i).Visible = 2
        End If
    Next i
    
    
    MinerWkBk.Activate
    Call M.LockUnlockWb(DB.WkBk, "lock", "DB")
    DB.WkBk.Save
End Sub

