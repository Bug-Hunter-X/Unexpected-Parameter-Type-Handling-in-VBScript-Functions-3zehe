Function MyFunc(param1)
  If IsEmpty(param1) Then
    ' Handle empty parameter
  ElseIf IsNumeric(param1) Then
    ' Process numeric parameter
  ElseIf IsDate(param1) Then
    ' Process Date parameter
  Else
    ' Handle other parameter types
  End If
End Function