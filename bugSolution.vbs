Function GetObjectSafe(progID)
  On Error Resume Next
  Set obj = GetObject(progID)
  If Err.Number <> 0 Then
    Err.Clear
    Set obj = CreateObject(progID)
    If Err.Number <> 0 Then
      ' Handle the case where CreateObject also fails
      MsgBox "Error: Could not create or get object: " & progID, vbCritical
      Set obj = Nothing
    End If
  End If
  Set GetObjectSafe = obj
End Function

' Example usage:
Set excelApp = GetObjectSafe("Excel.Application")
If Not excelApp Is Nothing Then
  MsgBox "Excel Application accessed successfully!"
  excelApp.Quit
  Set excelApp = Nothing
Else
  MsgBox "Could not access Excel Application."
End If