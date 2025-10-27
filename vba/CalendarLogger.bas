Attribute VB_Name = "CalendarLogger"
Option Explicit

' Minimal logger shim so the shared calendar worksheet code compiles. Swap this
' stub out for the project-specific logging engine that tracks grid edits.
Public Sub HandleGridChange(ByVal Target As Range)
    ' No-op placeholder
End Sub
