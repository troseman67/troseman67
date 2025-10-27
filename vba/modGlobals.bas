Attribute VB_Name = "modGlobals"
Option Explicit

' ----- Sheet names -----
Public Const S_CAL   As String = "Calendar Checklist"
Public Const S_DATA  As String = "Data"
Public Const S_GAUGE As String = "Gauge_Readings"

' ----- Calendar grid layout -----
Public Const FIRST_TASK_ROW As Long = 7
Public Const LAST_TASK_ROW  As Long = 24
Public Const ROW_DAYNUM     As Long = 6
Public Const FIRST_DAY_COL  As Long = 3
Public Const LAST_DAY_COL   As Long = 33

' ----- Banner / progress cells -----
Public Const COL_PROGRESS        As Long = 34
Public Const CELL_OVERALLPROG    As String = "L2"
Public Const CELL_LASTUPDATED    As String = "I2"
Public Const CELL_LASTUPDATEDBY  As String = "J3"

' ----- Header inputs -----
Public Const CELL_PLANT As String = "A5"
Public Const CELL_MONTH As String = "A3"
Public Const CELL_YEAR  As String = "B3"

' ----- Frequency codes -----
Public Const TF_DAILY   As Long = 1
Public Const TF_WEEKLY  As Long = 2
Public Const TF_MONTHLY As Long = 3

Public Function MonthKey(ByVal v As Variant) As Integer
    On Error GoTo Bad

    If IsDate(v) Then
        MonthKey = Month(CDate(v))
        Exit Function
    End If

    If IsNumeric(v) Then
        MonthKey = CInt(Val(v))
        If MonthKey >= 1 And MonthKey <= 12 Then Exit Function
    End If

    MonthKey = Month(CDate("1 " & Trim$(CStr(v)) & " 2024"))
    Exit Function

Bad:
    MonthKey = 0
End Function

Public Function NzStr(ByVal v As Variant) As String
    NzStr = Trim$(CStr(v))
End Function

Public Function FindDayHeaderRow(ws As Worksheet) As Long
    Dim r As Long, c As Long, cnt As Long, big As Long, v

    For r = FIRST_TASK_ROW - 1 To 3 Step -1
        cnt = 0: big = 0
        For c = FIRST_DAY_COL To LAST_DAY_COL
            v = ws.Cells(r, c).Value
            If IsNumeric(v) Then
                If v >= 1 And v <= 31 Then
                    cnt = cnt + 1
                    If v >= 10 Then big = big + 1
                End If
            End If
        Next c
        If cnt >= 10 And big >= 1 Then
            FindDayHeaderRow = r
            Exit Function
        End If
    Next r

    For r = 3 To 7
        cnt = 0: big = 0
        For c = FIRST_DAY_COL To LAST_DAY_COL
            v = ws.Cells(r, c).Value
            If IsNumeric(v) Then
                If v >= 1 And v <= 31 Then
                    cnt = cnt + 1
                    If v >= 10 Then big = big + 1
                End If
            End If
        Next c
        If cnt >= 10 And big >= 1 Then
            FindDayHeaderRow = r
            Exit Function
        End If
    Next r

    FindDayHeaderRow = ROW_DAYNUM
End Function
