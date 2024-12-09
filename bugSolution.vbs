Function CheckInput(input)
  If IsNumeric(input) Then
    MsgBox "Numeric Input: " & input
  ElseIf IsDate(input) Then
    MsgBox "Date Input: " & input
  Else
    MsgBox "String Input: " & input
  End If
End Function

'Instead of attempting overloading:
'Function CheckInput(input)
'  'Numeric handling'
'End Function
'
'Function CheckInput(input)
'  'String handling'
'End Function