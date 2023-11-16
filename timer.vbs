Dim counter
counter = InputBox("Enter the number of seconds for the countdown:", "Countdown Timer")
Do Until counter = 0
    WScript.Sleep 1000 ' Wait for 1 second
    counter = counter - 1
Loop
MsgBox "Time's up!", 64, "Countdown Timer"
