Dim name, message
name = InputBox("What's your crush's name?", "Crush")
message = "Dear " & name & "," & vbCrLf & vbCrLf & _
        "Whenever I see you, my heart skips a beat. Your smile lights up my world and your presence makes everything better. " & _
        "You are truly special and I'm grateful to have you in my life." & vbCrLf & vbCrLf & _
        "Yours," & vbCrLf & _
        "Secret Admirer"
MsgBox message, 64, "Message to Crush"
