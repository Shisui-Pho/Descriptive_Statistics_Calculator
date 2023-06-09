Attribute VB_Name = "FrequencyTable"
Option Explicit
Public Sub SetTable()
    Dim dRange As Double
    Dim iInterval As Integer
    Dim iWidth As Integer
    
    
    dRange = RangeValue
    Worksheets("Frequency table").Cells(5, 10).value = dRange
    
    iInterval = NIntervals
    Worksheets("Frequency table").Cells(6, 10).value = iInterval
    
    iWidth = ClassWidth(dRange, iInterval)
    Worksheets("Frequency table").Cells(7, 10).value = iWidth
    Call ClassC(iInterval, iWidth)
    Call MidPoint(iWidth, iInterval)
    Call CumulativeFrequency(iInterval)
    Call FM(iInterval)
    Call FM2(iInterval)
    Call Percent(iInterval)
End Sub
Private Sub ClassC(iInterval As Integer, iWidth As Integer)
    Dim dFirst As Integer
    Dim i As Integer
    Dim asData() As Variant
    Dim asInt() As Variant 'Intervals
    Dim s As String
    Dim asFreq() As Variant
    Dim value As Variant
    
    
    'Get the first value
    dFirst = Worksheets("Data entry").Cells(2, 3).value
    
    'Activate the second table
    Worksheets("Frequency table").Activate
    Cells(2, 1).Activate
    
    For i = 0 To iInterval - 1
        dFirst = dFirst + iWidth
        
        ActiveCell.Offset(i, 0).value = dFirst - 0.1
    Next i
    asData() = Range(SelectData)
    
     Worksheets("Frequency table").Activate
    s = "A2:A" & (1 + iInterval)
    
    asInt() = Range(s)
    
    asFreq() = Application.WorksheetFunction.Frequency(asData, asInt)
    Cells(2, 3).Activate
    i = 0
     'Application.WorksheetFunction.Frequency()
    For Each value In asFreq
        ActiveCell.Offset(i, 0).value = value
        i = i + 1
    Next
    ActiveCell.Offset(i - 1, 0).value = ""
End Sub
Private Sub MidPoint(iWidth As Integer, iInterval As Integer)
    Dim iLower As Integer
    Dim asInt() As Variant
    Dim s As String
    Dim value As Variant
    Dim i As Integer
    Dim iUpper As Integer
    
    Worksheets("Frequency table").Activate
    Cells(2, 1).Activate
    
    s = "A2:A" & (1 + iInterval)
    asInt() = Range(s)
    
    i = 0
    For Each value In asInt
        
        iLower = value - iWidth
        iUpper = value
        
        ActiveCell.Offset(i, 0).value = "[" & iLower & " , " & iUpper & ")"
        ActiveCell.Offset(i, 1).value = (iLower + iUpper) / 2
        i = i + 1
    Next
End Sub
Private Sub CumulativeFrequency(iInterval As Integer)
    Dim iValue As Integer
    Dim i As Integer

    Worksheets("Frequency table").Activate
    Cells(2, 3).Activate
    
    iValue = ActiveCell.Offset(0, 0).value
    
    For i = 0 To iInterval - 1
        ActiveCell.Offset(i, 1).value = iValue
        iValue = iValue + ActiveCell.Offset(i + 1, 0).value
    Next i
End Sub
Private Sub FM(iInterval As Integer)
    Dim i As Integer
    
    Worksheets("Frequency table").Activate
    Cells(2, 2).Activate
    
    For i = 0 To iInterval - 1
        ActiveCell.Offset(i, 3).value = Products(ActiveCell.Offset(i, 0).value, ActiveCell.Offset(i, 1).value, 1)
    Next i
End Sub
Private Sub FM2(iInterval As Integer)
        Dim i As Integer
    
    Worksheets("Frequency table").Activate
    Cells(2, 2).Activate
    
    For i = 0 To iInterval - 1
        ActiveCell.Offset(i, 4).value = Products(ActiveCell.Offset(i, 0).value, ActiveCell.Offset(i, 1).value, 2)
    Next i
End Sub
Private Sub Percent(iInterval)
    Dim i As Integer
    
    Worksheets("Frequency table").Activate
    Cells(2, 3).Activate
    
    For i = 0 To iInterval - 1
        ActiveCell.Offset(i, 4).value = (ActiveCell.Offset(i, 0).value / Cells(iInterval + 1, 4).value)
    Next i
End Sub
