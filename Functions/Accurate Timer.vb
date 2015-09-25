Declare Function getFrequency Lib "kernel32" Alias "QueryPerformanceFrequency" (cyFrequency As Currency) As Long
Declare Function getTickCount Lib "kernel32" Alias "QueryPerformanceCounter" (cyTickCount As Currency) As Long
Sub Try()
  Dim frequency, startTime, endTime As Currency: Dim result As Double
    
  'Start Timer
  getTickCount startTime
    
  'Stop Timer
  getTickCount endTime
  Trash = getFrequency(frequency)
    
  result = (endTime - startTime) / frequency
  MsgBox result
    
End Sub
