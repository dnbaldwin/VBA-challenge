Sub StockWorkSheets()



Dim currentStockName As String
Dim priceAtOpen As Double
Dim priceAtClose As Double
Dim currentStockVolumeTotal As Double
Dim previousStockName As String
Dim companyCounter As Double
Dim maxStockGain As Double
Dim maxStockGainTicker As String
Dim maxStockLoss As Double
Dim maxStockLossTicker As String
Dim maxStockVolumeTotal As Double
Dim maxStockVolumeTotalTicker As String


For Each ws In Worksheets

    previousStockName = ""
    currentStockName = ""
    nextStockName = ""
    
    maxStockGain = 0
    maxStockLoss = 0
    maxStockVolumeTotal = 0
    
    
    currentStockVolumeTotal = 0
    currentStockVolume = 0
    companyCounter = 0
    
    
    ' I have not created additional columns in this macro, but to do so would use the following:
    
    ' ws.Range("K1").EntireColumn.Insert
    
    
    ' Clear Data Area of Worksheet
    
        ws.Range("K1:T1").EntireColumn.Clear
        ws.Range("K1:T1").ColumnWidth = 16
        
    
    ' setup headers
    
        ws.Range("K1").Value = "Ticker"
        ws.Range("L1").Value = "Opening Price"
        ws.Range("M1").Value = "Closing Price"
        ws.Range("N1").Value = "Yearly Change"
        ws.Range("O1").Value = "Percent Change"
        ws.Range("P1").Value = "Total Stock Volume"
    
    ' find last row
    
        lastRow = Cells(Rows.Count, 1).End(xlUp).Row
        
    ' run the loop
    
        For currentRow = 2 To lastRow
        
            ' get data from each index in row
        
            currentStockName = ws.Cells(currentRow, 1).Value
            
            currentStockVolume = ws.Cells(currentRow, 7).Value
            
            previousStockName = ws.Cells(currentRow - 1, 1).Value
            
            nextStockName = ws.Cells(currentRow + 1, 1).Value
            
            currentStockVolumeTotal = currentStockVolumeTotal + currentStockVolume
            
            
            ' check if the company has changed
            
            If currentStockName <> previousStockName Then       ' the company has changed
            
                
                currentStockPriceAtOpen = ws.Cells(currentRow, 3)  ' The opening price is only the first row of new company
                
                
                ' reset company dependant variables
                
                companyCounter = companyCounter + 1             ' companyCounter will be the table row (+1)
                
                ' currentStockVolumeTotal = 0
                
            End If
            
            ' if the next stock is different, this is the last. use for
            ' getting closing price and summary data for current company
            
            If currentStockName <> nextStockName Then
            
                currentStockPriceAtClose = ws.Cells(currentRow, 6).Value
                
                ' calculate change in price for current stock
                
                currentStockPriceChange = currentStockPriceAtClose - currentStockPriceAtOpen
                
                currentStockPriceChangePerc = (currentStockPriceChange / currentStockPriceAtOpen)
                
                ws.Cells(companyCounter + 1, 11).Value = currentStockName
                ws.Cells(companyCounter + 1, 12).Value = currentStockPriceAtOpen
                ws.Cells(companyCounter + 1, 13).Value = currentStockPriceAtClose
                ws.Cells(companyCounter + 1, 14).Value = currentStockPriceChange
                ws.Cells(companyCounter + 1, 15).Value = FormatPercent(currentStockPriceChangePerc, 2)
                ws.Cells(companyCounter + 1, 16).Value = currentStockVolumeTotal
                
                ' check if this stock's gain, loss or volume is maximum
                ' if it is save it in max variable for later
                
                
                If currentStockPriceChangePerc > maxStockGain Then
                
                    maxStockGain = currentStockPriceChangePerc
                    maxStockGainTicker = currentStockName
                    
                ElseIf currentStockPriceChangePerc < maxStockLoss Then
                
                    maxStockLoss = currentStockPriceChangePerc
                    maxStockLossTicker = currentStockName
                
                End If
                
                If currentStockVolumeTotal > maxStockVolumeTotal Then
                
                    maxStockVolumeTotal = currentStockVolumeTotal
                    maxStockVolumeTotalTicker = currentStockName
                    
                End If
                
                ' set color format according to price change
                
                If currentStockPriceChange > 0 Then
                
                    ws.Cells(companyCounter + 1, 14).Interior.ColorIndex = 4
                    ws.Cells(companyCounter + 1, 15).Interior.ColorIndex = 4
                
                ElseIf currentStockPriceChange < 0 Then
                
                    ws.Cells(companyCounter + 1, 14).Interior.ColorIndex = 3
                    ws.Cells(companyCounter + 1, 15).Interior.ColorIndex = 3
                
                ElseIf currentStockPriceChange = 0 Then
                
                    ws.Cells(companyCounter + 1, 14).Interior.ColorIndex = 33
                    ws.Cells(companyCounter + 1, 15).Interior.ColorIndex = 33
                
                End If
                
                ' reset relevant variables
                currentStockVolumeTotal = 0
                
            End If
                  
        Next currentRow
        
        '------------------------------------------
        
        ' Analyse Summary Data
        
        ' set headers
        
        ws.Range("R1").Value = "Parameter"
        ws.Range("S1").Value = "Ticker"
        ws.Range("T1").Value = "Value"
        
        ws.Range("R2").Value = "Greatest % Increase"
        ws.Range("R3").Value = "Greatest % Decrease"
        ws.Range("R4").Value = "Greatest Total Volume"
        
        
        ' get ticker for each of the parameters
        
        ws.Range("S2").Value = maxStockGainTicker
        ws.Range("S3").Value = maxStockLossTicker
        ws.Range("S4").Value = maxStockVolumeTotalTicker
        
        ' get the values and format them for easier viewing (remove Sci.Not)
        
        ws.Range("T2").Value = FormatPercent(maxStockGain, 2)
        ws.Range("T3").Value = FormatPercent(maxStockLoss, 2)
        ws.Range("T4").Value = maxStockVolumeTotal
        ws.Range("T4").NumberFormat = "0"

    Next ws

End Sub
