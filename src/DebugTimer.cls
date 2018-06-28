VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DebugTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private Declare Function GetTickCount Lib "kernel32" () As Long

Private pStartTime As Long
Private pIntervalTime As Long
Private pIntervals As Collection

Private Sub Class_Initialize()
    Set pIntervals = New Collection
End Sub

Sub setStartTime()
    pStartTime = GetTickCount
    pIntervalTime = pStartTime
    pIntervals.Add pStartTime
End Sub

Function reportInterval() As Long
    pIntervals.Add GetTickCount - pStartTime
    reportInterval = maxInterval
    pIntervalTime = maxInterval
End Function

Function reportElapsed() As Long
    reportElapsed = GetTickCount - pStartTime
End Function



Private Function maxInterval() As Long
    maxInterval = pIntervals(pIntervals.Count)
End Function
