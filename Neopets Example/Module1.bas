Attribute VB_Name = "Module1"
Sub Pause(TimePaused)
    Dim Current
        Current = Timer
    Do While Timer - Current < Val(TimePaused)
        DoEvents
    Loop
End Sub
