Attribute VB_Name = "modWorkbookUtilities"
Option Explicit

' Ensures the data sheet used for calendar logging exists. Creates it if missing
' and applies a simple header row so downstream routines can rely on the layout.
Public Sub EnsureDataSheet()
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(S_DATA)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = S_DATA
        ws.Range("A1:J1").Value = Array("Plant", "Year", "Month", "Day", "TaskRow", "Slot", _
                                           "Value", "Comment", "EditedBy", "EditedAt")
    End If
End Sub

' Adds a refresh button to the calendar sheet if one is not already present. The
' button simply re-runs the calendar engine routine that rebuilds the visible
' grid for the currently selected plant/month/year.
Public Sub AddRefreshButton()
    Const BTN_NAME As String = "btnRefreshCalendar"

    Dim ws As Worksheet
    Dim shp As Shape

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(S_CAL)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    On Error Resume Next
    Set shp = ws.Shapes(BTN_NAME)
    On Error GoTo 0

    If shp Is Nothing Then
        Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, ws.Range("J2").Left, _
                                     ws.Range("J2").Top, 110, ws.Range("J2").Height + 6)
        With shp
            .Name = BTN_NAME
            .TextFrame.Characters.Text = "Refresh Calendar"
            .OnAction = "modCalendarEngine.LoadSelectedMonth"
        End With
    End If
End Sub
