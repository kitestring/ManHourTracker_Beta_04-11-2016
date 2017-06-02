VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub cmdLaunchQuery_Click()
    Dim UserEntry As String
    Dim MinerWkBk As Workbook
    Set MinerWkBk = ActiveWorkbook
    
    UserEntry = PW
    
    If UserEntry = "True" Then
        Application.ScreenUpdating = False
        Dim M As New Methods
        Call M.LockUnlockWb(MinerWkBk, "unlock", "Control")
        Load frmQuery
        frmQuery.Show
    ElseIf UserEntry = "False" Then
        MsgBox "Incorrect Password", vbCritical, "ACCESS DENIED"
        Exit Sub
    Else
        Sheets("TimeLogger").Select
        Exit Sub
    End If
    Call M.LockUnlockWb(MinerWkBk, "lock", "Control")
    Sheets("TimeLogger").Select
End Sub

Private Function PW() As String
    Dim M As New Methods
    Dim i As New UpdateInternational
    
    Dim user_value As String
    user_value = InputBox("Enter Password.", "PASSWORD")
    If user_value = "81643" Then
        PW = "True"
    ElseIf user_value = "kiteopen" Then
        Call M.LockUnlockWb(ActiveWorkbook, "unlock", "Control")
        PW = "Other"
    ElseIf user_value = "kiteclose" Then
        Call M.LockUnlockWb(ActiveWorkbook, "lock", "Control")
        PW = "Other"
    ElseIf user_value = "Intl" Then
        Call i.UpdateDB_International
        PW = "Other"
    Else
        PW = "False"
    End If
End Function

