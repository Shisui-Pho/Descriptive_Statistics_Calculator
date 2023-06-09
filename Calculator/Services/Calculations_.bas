Attribute VB_Name = "Calculations_"
Option Explicit

Public Function Mean() As Double
    Mean = Application.WorksheetFunction.Average(Range(SelectData()))
End Function
Public Function Varience(dMean As Double) As Double
    Dim iCount As Integer
    Dim dSum As Double
    Dim dBigSum As Double
    Dim dValue As Double
    Dim i As Integer
    Dim dCalculate As Double
    
    iCount = CountR
    Cells(2, 3).Activate
     For i = 0 To iCount
        dValue = ActiveCell.Offset(i, 0).value
        dValue = dValue * dValue
        
        dBigSum = dBigSum + dValue
    Next i
    
    dSum = Application.WorksheetFunction.Sum(Range(SelectData))
    
    iCount = iCount - 1
    dCalculate = 1 / (iCount - 1)
    dCalculate = dCalculate * (dBigSum - (1 / iCount) * (dSum * dSum))
    
    'MsgBox
    'MsgBox dBigSum
    Varience = dCalculate
End Function
Public Function Std(dVarience As Double) As Double
    Std = Application.WorksheetFunction.Power(dVarience, 0.5)
End Function
Public Function RangeValue() As Double
    Dim dSmallest As Double
    Dim dLargest As Double
    Dim iCount As Integer
    
    iCount = CountR
    
    'Cells(2, 3).Activate
    dSmallest = Cells(2, 3).value 'ActiveCell.Offset(0, 0).Value
    dLargest = Cells(iCount, 3).value 'ActiveCell.Offset(iCount - 1, 0).Value
     
    RangeValue = dLargest - dSmallest
    
End Function
Public Function Products(iMidpoint As Double, iFreq As Double, iPower As Integer) As Double
    Products = Application.WorksheetFunction.Power(iMidpoint, iPower) * iFreq
End Function

Public Function NIntervals() As Integer
    Dim iCount As Integer
    Dim iValue As Integer
    iCount = CountR
    
    iValue = 1 + (3.3 * Application.WorksheetFunction.Log10(iCount))
    NIntervals = iValue + 0.5
End Function

Public Function ClassWidth(dRange As Double, iInterval As Integer) As Integer
    ClassWidth = (dRange / iInterval) + 0.5
End Function


Public Function SelectData() As String
    Dim iCol As Integer
    
    iCol = CountR
    SelectData = "C2:C" & iCol
End Function
Public Function CountR() As Integer
    Dim iCol As Integer
    Worksheets("Data entry").Activate
    CountR = Application.WorksheetFunction.CountA(Range("B:B"))
End Function

