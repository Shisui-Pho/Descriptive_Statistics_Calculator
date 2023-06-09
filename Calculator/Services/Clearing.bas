Attribute VB_Name = "Clearing"
Option Explicit

Public Sub ClearEntry_1_2()
    Range("A2:C100").ClearContents
    Range("F9:F11").ClearContents
    Worksheets("Frequency table").Range("A2:G19").ClearContents
    Worksheets("Frequency table").Range("J5:J7").ClearContents
    Cells(1, 1).Activate
End Sub
