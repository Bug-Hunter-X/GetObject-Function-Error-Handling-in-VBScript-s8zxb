Function GetObject(progID) 
 On Error Resume Next 
 Set obj = GetObject(progID) 
 If Err.Number <> 0 Then 
 Err.Clear 
 Set obj = CreateObject(progID) 
 End If 
 Set GetObject = obj 
 End Function