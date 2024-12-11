Function MyFunc(param1)
  On Error Resume Next  ' Handle potential errors gracefully
  If IsEmpty(param1) Then
    Err.Clear
    ' Handle empty parameter: return a default value or perform alternative logic.
    MyFunc = 0 'Example default value
  ElseIf IsNumeric(param1) Then
    MyFunc = param1 * 2 'Process numeric parameter
  ElseIf IsDate(param1) Then
    MyFunc = CStr(param1) ' Process Date parameter
  ElseIf VarType(param1) = vbString Then
    MyFunc = UCase(param1) 'Handle String type specifically
  Else
    Err.Raise vbObjectError + 1, "MyFunc", "Unsupported parameter type: " & VarType(param1)
  End If
  On Error GoTo 0
End Function