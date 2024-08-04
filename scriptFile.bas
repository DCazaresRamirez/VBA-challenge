Attribute VB_Name = "Module1"
Sub QuarterlyStockAnalysis()
    Dim ws As Worksheet
    Dim lastRow As Long, outputRow As Long
    Dim ticker As String
    Dim firstOpen As Double, lastClose As Double
    Dim qChange As Double, qChangePercent As Double
    Dim totalVolume As Double
    Dim currentRow As Long
    
    ' Loop through each sheet
    For Each ws In ThisWorkbook.Sheets
        If ws.Name <> "Summary" Then
            lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
            
            ' Store the first open, last close, and total volume for each ticker
            Dim tickerDict As Object
            Set tickerDict = CreateObject("Scripting.Dictionary")
            
            ' Loop through the rows in the sheet
            For currentRow = 2 To lastRow
                ticker = ws.Cells(currentRow, 1).Value
                If Not tickerDict.exists(ticker) Then
                    ' New ticker, add to dictionary with first open, last close, and volume
                    tickerDict.Add ticker, Array(ws.Cells(currentRow, 3).Value, ws.Cells(currentRow, 6).Value, ws.Cells(currentRow, 7).Value)
                Else
                    ' Existing ticker, update last close and add to volume
                    tickerDict(ticker)(1) = ws.Cells(currentRow, 6).Value
                    tickerDict(ticker)(2) = tickerDict(ticker)(2) + ws.Cells(currentRow, 7).Value
                End If
            Next currentRow
            
            ' Add headers for new columns
            Dim lastColumn As Integer
            lastColumn = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
            ws.Cells(1, lastColumn + 1).Value = "Ticker"
            ws.Cells(1, lastColumn + 2).Value = "Quarterly Change"
            ws.Cells(1, lastColumn + 3).Value = "Percent Change"
            ws.Cells(1, lastColumn + 4).Value = "Total Stock Volume"
            
            ' Initialize variables to track greatest values
            Dim greatestIncreaseTicker As String
            Dim greatestIncreaseValue As Double
            Dim greatestDecreaseTicker As String
            Dim greatestDecreaseValue As Double
            Dim greatestVolumeTicker As String
            Dim greatestVolumeValue As Double

            greatestIncreaseValue = -1E+30 ' Initialize to a very low value
            greatestDecreaseValue = 1E+30  ' Initialize to a very high value
            greatestVolumeValue = 0        ' Initialize to 0

            ' Output results for each ticker
            Dim outputColStart As Integer
            outputColStart = lastColumn + 1

            Dim key As Variant
            outputRow = 2 ' Start from row 2 for data output
            
            For Each key In tickerDict.keys
                firstOpen = tickerDict(key)(0)
                lastClose = tickerDict(key)(1)
                totalVolume = tickerDict(key)(2)
                
                qChange = lastClose - firstOpen
                If firstOpen <> 0 Then
                    qChangePercent = (qChange / firstOpen) * 100
                Else
                    qChangePercent = 0
                End If
                
                ' Output results
                ws.Cells(outputRow, lastColumn + 1).Value = key
                ws.Cells(outputRow, lastColumn + 2).Value = qChange
                ws.Cells(outputRow, lastColumn + 3).Value = qChangePercent
                ws.Cells(outputRow, lastColumn + 4).Value = totalVolume
                
                ' Update ticker column
                ws.Cells(outputRow, 1).Value = key
                
                ' Check for greatest values
                If qChangePercent > greatestIncreaseValue Then
                    greatestIncreaseValue = qChangePercent
                    greatestIncreaseTicker = key
                End If
                
                If qChangePercent < greatestDecreaseValue Then
                    greatestDecreaseValue = qChangePercent
                    greatestDecreaseTicker = key
                End If
                
                If totalVolume > greatestVolumeValue Then
                    greatestVolumeValue = totalVolume
                    greatestVolumeTicker = key
                End If
                
                outputRow = outputRow + 1 ' Move to the next row
            Next key
            
            ' Add empty columns before greatest values summary
            ws.Cells(1, lastColumn + 5).Value = " "
            ws.Cells(1, lastColumn + 6).Value = "Metric"
            ws.Cells(1, lastColumn + 7).Value = "Ticker"
            ws.Cells(1, lastColumn + 8).Value = "Value"
            
            ' Output the greatest values with proper spacing
            ws.Cells(2, lastColumn + 6).Value = "Greatest % Increase"
            ws.Cells(2, lastColumn + 7).Value = greatestIncreaseTicker
            ws.Cells(2, lastColumn + 8).Value = greatestIncreaseValue
            
            ws.Cells(3, lastColumn + 6).Value = "Greatest % Decrease"
            ws.Cells(3, lastColumn + 7).Value = greatestDecreaseTicker
            ws.Cells(3, lastColumn + 8).Value = greatestDecreaseValue
            
            ws.Cells(4, lastColumn + 6).Value = "Greatest Total Volume"
            ws.Cells(4, lastColumn + 7).Value = greatestVolumeTicker
            ws.Cells(4, lastColumn + 8).Value = greatestVolumeValue
            
            ' Autofit columns for better visibility
            ws.Columns.AutoFit
        End If
    Next ws
End Sub
