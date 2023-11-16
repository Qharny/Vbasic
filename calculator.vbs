Dim num1, num2, operation, result
num1 = InputBox("Enter the first number:", "Calculator")
num2 = InputBox("Enter the second number:", "Calculator")
operation = InputBox("Enter the operation (+, -, *, /):", "Calculator")
If operation = "+" Then
    result = num1 + num2
ElseIf operation = "-" Then
    result = num1 - num2
ElseIf operation = "*" Then
    result = num1 * num2
ElseIf operation = "/" Then
    result = num1 / num2
Else
    MsgBox "Invalid operation!", 16, "Error"
    Exit Sub
End If
MsgBox "The result is: " & result, 64, "Calculator"
