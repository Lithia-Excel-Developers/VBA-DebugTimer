Attribute VB_Name = "test_DebugTimer"
Option Explicit


Sub test()
    
    'one way to make a new object
    Dim dbTimer1 As New DebugTimer
    
    'another way to make a new object
    Dim dbTimer2 As DebugTimer
    Set dbTimer2 = New DebugTimer
    
    
    dbTimer1.setStartTime
    
    'create a delay
    Dim i As Long
    Dim j As String
    For i = 0 To 300000
        j = Cells(1, 1).Value
    Next i
    
    dbTimer2.setStartTime
    
    
    Debug.Print "--------------------------------" & vbCrLf & vbCrLf
    Debug.Print "T1 Interval " & dbTimer1.reportInterval
    
    'create a second delay
    For i = 0 To 600000
        j = Cells(1, 1).Value
    Next i

    Debug.Print "T1 Interval " & dbTimer1.reportInterval
    Debug.Print "T2 Interval " & dbTimer2.reportInterval
    
    Debug.Print "T1 elapsed " & dbTimer1.reportElapsed
    Debug.Print "T2 elapsed " & dbTimer2.reportElapsed
    
End Sub




