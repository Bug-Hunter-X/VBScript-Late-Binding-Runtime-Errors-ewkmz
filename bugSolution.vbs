On Error Resume Next
Set obj = CreateObject("Some.Object")
If Err.Number <> 0 Then
  MsgBox "Error creating or accessing object: " & Err.Description
  Err.Clear
Else
  ' Access object members safely here
  If TypeName(obj) = "SomeObjectType" Then
    ' obj.SomeMethod
  Else
    MsgBox "Object is not the expected type."
  End If
End If
Set obj = Nothing
On Error GoTo 0