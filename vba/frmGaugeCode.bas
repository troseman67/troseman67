Attribute VB_Name = "frmGaugeCode"
Option Explicit

' Code-behind for frmGauge userform. Import this into the form module.

Private mBound As Boolean

Public plant As String
Public yr As Long
Public monText As String
Public DayNum As Long
Public TaskRow As Long
Public TargetCol As Long

Private Sub UserForm_Initialize()
    Me.Caption = "Enter Gauge Readings (3×/day)"
    mBound = False
    On Error Resume Next
    lblHdr.Caption = vbNullString
    On Error GoTo 0
End Sub

Public Sub BindContext()
    lblHdr.Caption = plant & " — " & monText & " " & DayNum & ", " & yr
    PrefillReadings
    mBound = True
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim a As String
    Dim b As String
    Dim c As String
    Dim note As String

    a = Trim$(txtAM.Text)
    b = Trim$(txtMID.Text)
    c = Trim$(txtPM.Text)
    note = Trim$(txtComment.Text)

    If Len(a) > 0 And Not IsNumeric(a) Then GoTo Bad
    If Len(b) > 0 And Not IsNumeric(b) Then GoTo Bad
    If Len(c) > 0 And Not IsNumeric(c) Then GoTo Bad

    DoSaveGaugeReadings a, b, c, note
    Exit Sub

Bad:
    MsgBox "Please enter numbers only for AM/MID/PM (or leave blank).", vbExclamation
End Sub

Private Sub EnsureGaugeLog(ByRef ws As Worksheet)
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(S_GAUGE)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = S_GAUGE
        ws.Range("A1:J1").Value = Array("Plant", "Year", "Month", "Day", "TaskRow", "Slot", "Reading", "Comment", "EditedBy", "EditedAt")
    End If
End Sub

Private Sub PrefillReadings()
    Dim ws As Worksheet
    Dim lr As Long
    Dim r As Long
    Dim mkForm As Integer
    Dim mkRow As Integer
    Dim p As String
    Dim m As String
    Dim y As Long
    Dim d As Long
    Dim tr As Long

    EnsureGaugeLog ws

    p = UCase$(plant)
    m = monText
    y = yr
    d = DayNum
    tr = TaskRow
    mkForm = MonthKey(m)
    lr = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    txtAM.Text = vbNullString
    txtMID.Text = vbNullString
    txtPM.Text = vbNullString
    txtComment.Text = vbNullString

    For r = 2 To lr
        If UCase$(CStr(ws.Cells(r, 1).Value)) = p _
           And CLng(Val(ws.Cells(r, 2).Value)) = y _
           And CLng(Val(ws.Cells(r, 4).Value)) = d _
           And CLng(Val(ws.Cells(r, 5).Value)) = tr Then

            mkRow = MonthKey(ws.Cells(r, 3).Value)
            If mkRow = mkForm Then
                Select Case UCase$(CStr(ws.Cells(r, 6).Value))
                    Case "AM":  If Len(txtAM.Text) = 0 Then txtAM.Text = CStr(ws.Cells(r, 7).Value)
                    Case "MID": If Len(txtMID.Text) = 0 Then txtMID.Text = CStr(ws.Cells(r, 7).Value)
                    Case "PM":  If Len(txtPM.Text) = 0 Then txtPM.Text = CStr(ws.Cells(r, 7).Value)
                End Select

                If Len(txtComment.Text) = 0 Then txtComment.Text = CStr(ws.Cells(r, 8).Value)
            End If
        End If
    Next r
End Sub

Private Sub DoSaveGaugeReadings(ByVal amText As String, _
                                ByVal midText As String, _
                                ByVal pmText As String, _
                                ByVal note As String)

    On Error GoTo CleanFail

    Dim wsL As Worksheet
    Dim wsC As Worksheet
    Dim lr As Long
    Dim r As Long
    Dim lastRow As Long
    Dim who As String
    Dim nowDt As Date
    Dim cnt As Long
    Dim existAM As Variant
    Dim existMID As Variant
    Dim existPM As Variant
    Dim existNote As String
    Dim finalAM As Variant
    Dim finalMID As Variant
    Dim finalPM As Variant
    Dim finalNote As String
    Dim mkForm As Integer
    Dim numVal As Double

    EnsureGaugeLog wsL
    Set wsC = ThisWorkbook.Worksheets(S_CAL)

    who = Environ$("Username")
    nowDt = Now

    existAM = Empty
    existMID = Empty
    existPM = Empty
    existNote = vbNullString
    mkForm = MonthKey(monText)

    lastRow = wsL.Cells(wsL.Rows.Count, "A").End(xlUp).Row
    For r = 2 To lastRow
        If UCase$(CStr(wsL.Cells(r, 1).Value)) = UCase$(plant) _
           And CLng(Val(wsL.Cells(r, 2).Value)) = yr _
           And CLng(Val(wsL.Cells(r, 4).Value)) = DayNum _
           And CLng(Val(wsL.Cells(r, 5).Value)) = TaskRow _
           And MonthKey(wsL.Cells(r, 3).Value) = mkForm Then

            Select Case UCase$(CStr(wsL.Cells(r, 6).Value))
                Case "AM":  existAM = wsL.Cells(r, 7).Value
                Case "MID": existMID = wsL.Cells(r, 7).Value
                Case "PM":  existPM = wsL.Cells(r, 7).Value
            End Select

            If Len(existNote) = 0 Then existNote = CStr(wsL.Cells(r, 8).Value)
        End If
    Next r

    finalAM = IIf(Len(amText) > 0 And amText <> "-", amText, existAM)
    finalMID = IIf(Len(midText) > 0 And midText <> "-", midText, existMID)
    finalPM = IIf(Len(pmText) > 0 And pmText <> "-", pmText, existPM)
    finalNote = IIf(Len(note) > 0, note, existNote)

    If amText = "-" Then finalAM = Empty
    If midText = "-" Then finalMID = Empty
    If pmText = "-" Then finalPM = Empty

    For r = lastRow To 2 Step -1
        If UCase$(CStr(wsL.Cells(r, 1).Value)) = UCase$(plant) _
           And CLng(Val(wsL.Cells(r, 2).Value)) = yr _
           And CLng(Val(wsL.Cells(r, 4).Value)) = DayNum _
           And CLng(Val(wsL.Cells(r, 5).Value)) = TaskRow _
           And MonthKey(wsL.Cells(r, 3).Value) = mkForm Then
            wsL.Rows(r).Delete
        End If
    Next r

    cnt = 0
    If Not IsEmpty(finalAM) And IsNumeric(finalAM) Then
        numVal = CDbl(finalAM)
        lr = wsL.Cells(wsL.Rows.Count, "A").End(xlUp).Row + 1
        wsL.Cells(lr, 1).Resize(1, 10).Value = Array(plant, yr, monText, DayNum, TaskRow, "AM", numVal, finalNote, who, nowDt)
        cnt = cnt + 1
    End If

    If Not IsEmpty(finalMID) And IsNumeric(finalMID) Then
        numVal = CDbl(finalMID)
        lr = wsL.Cells(wsL.Rows.Count, "A").End(xlUp).Row + 1
        wsL.Cells(lr, 1).Resize(1, 10).Value = Array(plant, yr, monText, DayNum, TaskRow, "MID", numVal, finalNote, who, nowDt)
        cnt = cnt + 1
    End If

    If Not IsEmpty(finalPM) And IsNumeric(finalPM) Then
        numVal = CDbl(finalPM)
        lr = wsL.Cells(wsL.Rows.Count, "A").End(xlUp).Row + 1
        wsL.Cells(lr, 1).Resize(1, 10).Value = Array(plant, yr, monText, DayNum, TaskRow, "PM", numVal, finalNote, who, nowDt)
        cnt = cnt + 1
    End If

    Application.EnableEvents = False
    wsC.Cells(TaskRow, TargetCol).Value = cnt
    Application.EnableEvents = True

    CalendarLogger.HandleGridChange wsC.Cells(TaskRow, TargetCol)
    Unload Me
    Exit Sub

CleanFail:
    Application.EnableEvents = True
    MsgBox "Unable to save gauge readings: " & Err.Description, vbExclamation
End Sub
