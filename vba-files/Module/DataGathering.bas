Attribute VB_Name = "DataGathering"
Option Explicit

' Function to get data from multiple sheets using a dictionary for standardization
Public Sub GatherData()
    Dim mainWS As Worksheet
    Dim i As Long
    Dim columnRanges As Variant

    Set mainWS = ThisWorkbook.Worksheets(1)

    For i = 2 To ThisWorkbook.Worksheets.Count
        columnRanges = GetColumnRanges(ThisWorkbook.Worksheets(i))
        ' TODO: consume columnRanges to populate master sheet
    Next i
End Sub

Private Function GetColumn(ws As Worksheets) As As Integer(4)
    Dim dict As Object
    Set dict = ProductDict_SetUp()

    
    

End Function