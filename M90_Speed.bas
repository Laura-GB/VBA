Attribute VB_Name = "M90_Speed"
Option Explicit

Sub SpeedUp()

    Dim Start As Date
    Dim Finish As Date
    Dim Counter As Long
    
    '=== Set up sheet
    Range("B1").Select
    Range("C1").Formula = "=B1 * 2"
    
    '=== Start Time
    Start = Now
    Debug.Print "Start", Start
    
    '=== Pointless loop that takes a while
    For Counter = 1 To 10000
        ActiveCell.Value = Counter
        ActiveCell.Offset(1, 0).Select
        ActiveCell.Offset(-1, 0).Select
    Next Counter

    '=== Finish Time
    Finish = Now
    Debug.Print "Finish", Finish
    Debug.Print "Time", (Finish - Start) * 24 * 60 * 60 & " seconds"
    Debug.Print "========================================"

End Sub
