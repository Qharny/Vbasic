Dim birthday
birthday = InputBox("Please enter your birthday (MM/DD):", "Birthday Surprise")

If birthday = DatePart("m", Date) & "/" & DatePart("d", Date) Then
    MsgBox "Happy Birthday! 🎉🎂🎈", 64, "Birthday Surprise"
Else
    MsgBox "It's not your birthday today, but when the day comes, remember that you're amazing! 😊", 64, "Birthday Surprise"
End If
