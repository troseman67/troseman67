VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Calendar_Checklist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mHandlingDbl As Boolean

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    If Target Is Nothing Then Exit Sub
    If Target.CountLarge > 1 Then Exit Sub

    On Error GoTo CleanFail

    Dim grid As Range
    Set grid = Me.Range(Me.Cells(FIRST_TASK_ROW, FIRST_DAY_COL), _
                        Me.Cells(LAST_TASK_ROW, LAST_DAY_COL))
    If Intersect(Target, grid) Is Nothing Then Exit Sub
    If Target.Row <> FIRST_TASK_ROW Then Exit Sub

    Cancel = True
    If mHandlingDbl Then Exit Sub

    mHandlingDbl = True
    Application.EnableEvents = False
    modGaugeForm.OpenGaugeFormForCell Target

CleanExit:
    Application.EnableEvents = True
    mHandlingDbl = False
    Exit Sub

CleanFail:
    MsgBox "Unable to open the gauge entry form: " & Err.Description, vbExclamation
    Resume CleanExit
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    If Target Is Nothing Then Exit Sub

    On Error GoTo CleanFail

    Dim key As Range
    Set key = Me.Range(CELL_PLANT & "," & CELL_MONTH & "," & CELL_YEAR)
    If Intersect(Target, key) Is Nothing Then Exit Sub

    Application.EnableEvents = False

    On Error Resume Next
    modCalendarEngine.LoadSelectedMonth
    On Error GoTo CleanFail

    Application.Calculate
    DoEvents

    On Error Resume Next
    modGridValidation.ApplyGridValidation
    modGaugeSync.RefreshGaugeGridFromLog
    On Error GoTo CleanFail

CleanExit:
    Application.EnableEvents = True
    Exit Sub

CleanFail:
    MsgBox "Calendar refresh failed: " & Err.Description, vbExclamation
    Resume CleanExit
End Sub
