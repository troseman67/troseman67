VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Workbook_Open()
    On Error GoTo CleanFail

    Application.EnableEvents = False
    EnsureDataSheet
    AddRefreshButton

CleanExit:
    Application.EnableEvents = True
    Exit Sub

CleanFail:
    MsgBox "An error occurred while initialising the workbook: " & Err.Description, vbExclamation
    Resume CleanExit
End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    On Error GoTo CleanExit

    If Sh.Name <> S_CAL Then Exit Sub
    If Intersect(Target, Sh.Range(CELL_MONTH & "," & CELL_YEAR & "," & CELL_PLANT)) Is Nothing Then Exit Sub

    modCalendarEngine.LoadSelectedMonth

CleanExit:
    If Err.Number <> 0 Then
        Debug.Print "Workbook_SheetChange error:", Err.Number, Err.Description
        Err.Clear
    End If
End Sub

Public Sub WiringFix()
    EnsureDataSheet
    Application.EnableEvents = True
End Sub
