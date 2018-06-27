VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DebugTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False



Private pStartTime As Date
Private pIntervalTime As Date

Sub setStartTime()
    pStartTime = Now
    pIntervalTime = Now
End Sub

Function reportInterval() As Double
    reportInterval = Now - pIntervalTime
    pIntervalTime = Now
End Function

Function reportElapsed() As Double
    reportElapsed = Now - pStartTime
End Function
