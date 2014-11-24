Attribute VB_Name = "Module1"
Public Sub Pause(Duration)
  Dim numTime
  numTime = Timer
  Do While Timer - numTime < Duration
    DoEvents
  Loop
End Sub
