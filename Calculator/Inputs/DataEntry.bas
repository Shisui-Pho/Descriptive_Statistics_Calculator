Attribute VB_Name = "DataEntry"
Option Explicit
Sub Main()
    Dim dMean As Double
    Dim dVarience As Double
    Dim dStd As Double
    
    Worksheets("Data entry").Activate
    
    Call DataEntry
    Call OrderedData
    
    
    If Cells(2, 3).value = "Unordered data" Then
        Call ClearEntry_1_2
        Exit Sub
    End If
    
    dMean = Mean()
    Cells(9, 6).value = dMean
    
    dVarience = Varience(dMean)
    Cells(10, 6).value = dVarience
    
    dStd = Std(dVarience)
    Cells(11, 6).value = dStd
    
    Call SetTable
    
    Worksheets("Data entry").Activate
    MsgBox "Go to Frequency spreadsheet to find the Frequency table", vbOKOnly + vbInformation, "COMPLETE"
End Sub

'Data entry
Private Sub DataEntry()

    'Variables
    Dim sInput As String
    Dim iCount As Double
    Cells(2, 2).Activate
    
    sInput = " "
    iCount = Application.WorksheetFunction.CountA(Range("B:B")) - 1
    
    'Get user input and do some error checking
    Do While StrPtr(sInput) <> 0
        sInput = InputBox("Enter values one-by-one", "Input values")
            
        If StrPtr(sInput) = 0 Then
            Exit Sub
        End If
        
        If IsNumeric(sInput) Then
            ActiveCell.Offset(iCount, 0).value = sInput 'For inputs
            ActiveCell.Offset(iCount, 0 - 1).value = iCount + 1 ' For Id's
        Else
            MsgBox "Please enter a numeric value", vbOKOnly + vbInformation, "NON-NUMERIC VALUE DETECTED"
            Call DataEntry
        End If
        iCount = iCount + 1
    Loop
     'DataEntry = iCount + 1
End Sub

'Order Data
Private Sub OrderedData()
    Dim iCol As Integer
    Dim i As Integer
    Dim vRawData()
    Dim s As String
    Dim bb As Variant
    
    
    Cells(4, 2).Activate
    iCol = Application.WorksheetFunction.CountA(Range("B:B"))
   s = "B2:B" & iCol
   vRawData() = Range(s)
    i = 0
    Cells(2, 3).Activate

    For Each bb In vRawData
        ActiveCell.Offset(i, 0).value = bb
        i = i + 1
    Next bb
    s = "C2:C" & iCol
    Range(s).Sort Key1:=Range(s), order1:=xlAscending, Header:=xlNo
    
End Sub
Private Sub DismissForm()
    Unload CfrmSplashScreen
End Sub
