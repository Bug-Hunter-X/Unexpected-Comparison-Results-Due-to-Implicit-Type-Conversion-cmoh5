Function f(a,b)
  If IsNumeric(a) And IsNumeric(b) Then 'Check if both inputs are numeric
    If CDbl(a) > CDbl(b) Then 
      MsgBox "a is greater than b"
    ElseIf CDbl(a) < CDbl(b) Then
      MsgBox "a is less than b"
    Else
      MsgBox "a is equal to b"
    End If
  Else
    MsgBox "Error: Inputs must be numbers"
  End If
end function

Call f(5, "10")
Call f(5, 10)