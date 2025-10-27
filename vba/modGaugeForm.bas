Attribute VB_Name = "modGaugeForm"
Option Explicit

Private mIsShowing As Boolean

Public Sub OpenGaugeFormForCell(ByVal Target As Range)
    If Target Is Nothing Then Exit Sub

    On Error GoTo Fail

    If mIsShowing Then Exit Sub
    mIsShowing = True

    Application.EnableEvents = False

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(S_CAL)

    Dim hdrRow As Long
    hdrRow = FindDayHeaderRow(ws)
    If hdrRow = 0 Then
        MsgBox "Day header row not found.", vbExclamation
        GoTo CleanExit
    End If

    Dim dn As Long
    dn = CoerceDayNum(ws.Cells(hdrRow, Target.Column).Value)
    If dn < 1 Or dn > 31 Then
        MsgBox "Day number out of range (1..31).", vbExclamation
        GoTo CleanExit
    End If

    Dim plant As String
    Dim mon As String
    Dim yr As Long

    plant = CStr(ws.Range(CELL_PLANT).Value)
    mon = CStr(ws.Range(CELL_MONTH).Value)
    yr = CLng(Val(ws.Range(CELL_YEAR).Value))

    If Len(plant) = 0 Or Len(mon) = 0 Or yr = 0 Then
        MsgBox "Select Plant (A5), Month (A3), and Year (B3) first.", vbExclamation
        GoTo CleanExit
    End If

    Unload frmGauge
    Load frmGauge

    With frmGauge
        .plant = plant
        .yr = yr
        .monText = mon
        .DayNum = dn
        .TaskRow = FIRST_TASK_ROW
        .TargetCol = Target.Column
        .BindContext
        .Show
    End With

CleanExit:
    Application.EnableEvents = True
    mIsShowing = False
    Exit Sub

Fail:
    Application.EnableEvents = True
    mIsShowing = False
    MsgBox "OpenGaugeFormForCell error: " & Err.Description, vbExclamation
End Sub

Private Function CoerceDayNum(ByVal v As Variant) As Long
    If IsNumeric(v) Then
        CoerceDayNum = CLng(Val(v))
        Exit Function
    End If

    If IsDate(v) Then
        CoerceDayNum = Day(CDate(v))
        Exit Function
    End If

    CoerceDayNum = CLng(Val(CStr(v)))
End Function
