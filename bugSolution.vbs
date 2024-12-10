Function f(a, b)
  If IsEmpty(a) Or a = "" Then
    a = 0
  ElseIf IsNumeric(a) = False Then
      Err.Raise vbError, , "Invalid input: a must be numeric or empty!" 
  End If
  If IsEmpty(b) Or b = "" Then
    b = 0
  ElseIf IsNumeric(b) = False Then
    Err.Raise vbError, , "Invalid input: b must be numeric or empty!" 
  End If
  c = CDbl(a) + CDbl(b)
  f = c
End Function

MsgBox f(1, "")
MsgBox f(1, "abc")
MsgBox f("","")
MsgBox f(1,2) 