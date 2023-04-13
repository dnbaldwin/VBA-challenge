Sub stocks2()


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



previousStockName = ""
currentStockName = ""

currentStockVolumeTotal = 0
currentStockVolume = 0
companyCounter = 0

' Clear Data Area of Worksheet

    Range("K1:T1").EntireColumn.Clear
    

' setup headers

    Range("K1").Value = "Ticker"
    Range("L1").Value = "Opening Price"
    Range("M1").Value = "Closing Price"
    Range("N1").Value = "Yearly Change"
    Range("O1").Value = "Percent Change"
    Range("P1").Value = "Total Stock Volume"

' find last row

    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
' run the loop

    For currentRow = 2 To lastRow
    
        ' get data from each index in row
    
        currentStockName = Cells(currentRow, 1).Value
        
        currentStockVolume = Cells(currentRow, 7).Value
        
        previousStockName = Cells(currentRow - 1, 1).Value
        
        nextStockName = Cells(currentRow + 1, 1).Value
        
        currentStockVolumeTotal = currentStockVolumeTotal + currentStockVolume
        
        
        ' check if the company has changed
        
        If currentStockName <> previousStockName Then       ' the company has changed
        
            
            currentStockPriceAtOpen = Cells(currentRow, 3)  ' The opening price is only the first row of new company
            
            
            ' reset company dependant variables
            
            companyCounter = companyCounter + 1             ' companyCounter will be the table row (+1)
            
            ' currentStockVolumeTotal = 0
            
            
        
        End If
        
        ' if the next stock is different, this is the last. use for
        ' getting closing price and summary data for current company
        
        
        If currentStockName <> nextStockName Then
        
            currentStockPriceAtClose = Cells(currentRow, 6).Value
            
            ' calculate change in price for current stock
            
            currentStockPriceChange = currentStockPriceAtClose - currentStockPriceAtOpen
            
            currentStockPriceChangePerc = (currentStockPriceChange / currentStockPriceAtOpen)
            
            Cells(companyCounter + 1, 11).Value = currentStockName
            Cells(companyCounter + 1, 12).Value = currentStockPriceAtOpen
            Cells(companyCounter + 1, 13).Value = currentStockPriceAtClose
            Cells(companyCounter + 1, 14).Value = currentStockPriceChange
            Cells(companyCounter + 1, 15).Value = FormatPercent(currentStockPriceChangePerc, 2)
            Cells(companyCounter + 1, 16).Value = currentStockVolumeTotal
            
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
            
                Cells(companyCounter + 1, 14).Interior.ColorIndex = 4
            
            ElseIf currentStockPriceChange < 0 Then
            
                Cells(companyCounter + 1, 14).Interior.ColorIndex = 3
            
            ElseIf currentStockPriceChange = 0 Then
            
                Cells(companyCounter + 1, 14).Interior.ColorIndex = 33
            
            End If
            
            ' reset relevant variables
            currentStockVolumeTotal = 0
            
        End If
              
    Next currentRow
    
    '------------------------------------------
    
    ' Analyse Summary Data
    
    ' set headers
    
    Range("R1").Value = "Parameter"
    Range("S1").Value = "Ticker"
    Range("T1").Value = "Value"
    
    Range("R2").Value = "Greatest % Increase"
    Range("R3").Value = "Greatest % Decrease"
    Range("R4").Value = "Greatest Total Volume"
    
    
    ' get ticker for each of the parameters
    
    Range("S2").Value = maxStockGainTicker
    Range("S3").Value = maxStockLossTicker
    Range("S4").Value = maxStockVolumeTotalTicker
    
    ' get the values and format them for easier viewing (remove Sci.Not)
    
    Range("T2").Value = FormatPercent(maxStockGain, 2)
    Range("T3").Value = FormatPercent(maxStockLoss, 2)
    Range("T4").Value = maxStockVolumeTotal
    Range("T4").NumberFormat = "0"


End Sub
