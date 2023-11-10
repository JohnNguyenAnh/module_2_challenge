# module_2_challenge

Sub module2challenge()
    Dim i As Long
    Dim j As Long
    Dim last_row As Long
    Dim ticker_row As Long
    Dim yearChange As Long
    Dim stock_name_row As Long
    Dim percentChange As Double
    Dim ws As Worksheet
    ws_num = ThisWorkbook.Worksheets.Count
    
    For j = 1 To ws_num
        ThisWorkbook.Worksheets(j).Activate
        'Initialize variables
        ticker_row = 2
        yearChange = 2
        stock_name_row = 2
        percent_change_row = 2
        sum_volume = 2
        
        'Initialize names for new column
        Range("k1").Value = "Ticker"
        Range("L1").Value = "Yearly Change"
        Range("M1").Value = "Percent Change"
        Range("N1").Value = "Total Stock Volume"
        Range("Q2").Value = "Greatest % increase"
        Range("Q3").Value = "Greatest total volume"
        Range("R1").Value = "Ticker"
        Range("S1").Value = "Value"
        
        'Find the last row of column A
        last_row = Range("A1").End(xlDown).Row
    
        For i = 2 To last_row
            If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            'Sort the ticker
            Cells(ticker_row, 11).Value = Cells(i, 1)
        
            'Find yearly change value
            open_price_firstday = Cells(stock_name_row, 3).Value
            close_price_lastday = Cells(i, 6).Value
            Cells(yearChange, 12).Value = (close_price_lastday - open_price_firstday)
        
            'Find percent change value
            Cells(percent_change_row, 13).Value = ((close_price_lastday - open_price_firstday) / open_price_firstday) * 100
        
            
            Cells(sum_volume, 14).Value = Application.WorksheetFunction.Sum(Cells(ticker_row, 7).Resize(i - ticker_row + 1, 1))
            
            
            
            yearChange = yearChange + 1
            ticker_row = ticker_row + 1
            stock_name_row = (i + 1)
            percent_change_row = percent_change_row + 1
            sum_volume = sum_volume + 1
        
            End If
        Next i
    Next
End Sub

Sub MaxMinTotalvolume()
    Dim i As Long
    Dim lastRow As Long
    
    
     
        'Find the last row of column K
        lastRow = Range("K2").End(xlDown).Row
        
        'Initialize variables
        maxValue = Cells(2, 13).Value
        minValue = Cells(2, 13).Value
        maxName = Cells(2, 11).Value
        minName = Cells(2, 11).Value
        maxVolumeName = Cells(2, 11).Value
        maxVolume = Cells(2, 14).Value
    
        For i = 2 To lastRow
            currentValue = Cells(i, 13).Value
            currentName = Cells(i, 11).Value
            currentVolume = Cells(i, 14).Value
        
            'Loop through data and looking for the max value
            If currentValue > maxValue Then
                maxValue = currentValue
                maxName = currentName
            End If
            'Loop through data and looking for the min value
            If currentValue < minValue Then
                minValue = currentValue
                minName = currentName
            End If
            'Loop through data and looking for the max volume value
            If currentVolume > maxVolume Then
                maxVolume = currentVolume
                maxVolumeName = currentName
            End If
        Next i
        'Assign value to appropriate column and row
        Cells(2, 19).Value = maxValue
        Cells(2, 18).Value = maxName
        Cells(3, 19).Value = minValue
        Cells(3, 18).Value = minName
        Cells(4, 18).Value = maxVolumeName
        Cells(4, 19).Value = maxVolume
    
End Sub



