VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Sub StockAnalysis()

'Make code run on all worksheets
Dim ws As Worksheet

For Each ws In Worksheets


'Set Variables
Dim Ticker As String
Dim PriceOpen As Double
Dim PriceHigh As Double
Dim PriceLow As Double
Dim PriceClose As Double
Dim TotalVolume As Double
Dim LastRow As Double
Dim StartRow As Double
Dim i As Double
Dim OpenRow As Double
Dim PriceChange As Double
Dim PercentChange As Double
Dim MaxIncreaseIndex As Double
Dim MaxDecreaseIndex As Double
Dim MaxVolumeIndex As Double

StartRow = 2
TotalVolume = 0
OpenRow = 2


'Set Last Row

LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Set Headers

ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Stock Volume"
ws.Cells(1, 17).Value = "Value"
ws.Cells(2, 15).Value = "Greatest%Increase"
ws.Cells(3, 15).Value = "Greatest%Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"


'Loop through Data

For i = 2 To LastRow
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        ws.Cells(StartRow, 9).Value = ws.Cells(i, 1).Value
        
        'Get open close values
        PriceClose = ws.Cells(i, 6).Value
        PriceOpen = ws.Cells(OpenRow, 3).Value
        'Calculate price changes
        PriceChange = PriceClose - PriceOpen
        PercentChange = PriceChange / PriceOpen
        
        'Input price change values into column
        ws.Cells(StartRow, 10).Value = PriceChange
        ws.Cells(StartRow, 11).Value = PercentChange
        'Formats price change to percentage
        ws.Cells(StartRow, 11).NumberFormat = "0.00%"
        'Add color to price change
        If ws.Cells(StartRow, 10).Value > 0 Then
            ws.Cells(StartRow, 10).Interior.ColorIndex = 4
        ElseIf ws.Cells(StartRow, 10).Value < 0 Then
            ws.Cells(StartRow, 10).Interior.ColorIndex = 3
        End If
        
        
        'solve for total volume
        TotalVolume = TotalVolume + ws.Cells(i, 7).Value
        ws.Cells(StartRow, 12).Value = TotalVolume
        
        'Goes to the next row
        StartRow = StartRow + 1
        TotalVolume = 0
        OpenRow = i + 1
        
    Else
        'input total volume into column
        TotalVolume = TotalVolume + ws.Cells(i, 7).Value
        
    End If
    
    Next i
    
    'create data summary
    ws.Cells(2, 17).Value = WorksheetFunction.Max(ws.Range("K2:K" & LastRow))
    ws.Cells(2, 17).NumberFormat = "0.00%"
    ws.Cells(3, 17).Value = WorksheetFunction.Min(ws.Range("K2:K" & LastRow))
    ws.Cells(3, 17).NumberFormat = "0.00%"
    ws.Cells(4, 17).Value = WorksheetFunction.Max(ws.Range("L2:L" & LastRow))
    MaxIncreaseIndex = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & LastRow)), ws.Range("K2:K" & LastRow), 0)
    MaxDecreaseIndex = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & LastRow)), ws.Range("K2:K" & LastRow), 0)
    MaxVolumeIndex = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & LastRow)), ws.Range("L2:L" & LastRow), 0)
    ws.Cells(2, 16).Value = ws.Cells(MaxIncreaseIndex + 1, 9)
    ws.Cells(3, 16).Value = ws.Cells(MaxDecreaseIndex + 1, 9)
    ws.Cells(4, 16).Value = ws.Cells(MaxVolumeIndex + 1, 9)
    


Next ws

End Sub
