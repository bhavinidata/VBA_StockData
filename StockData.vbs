Sub StockData()

Dim openingPriceChk As Bool
Dim totalVolumeTableRow, lastRow, openingPriceRow, maxTotalStockVolume, maxStockCell As Long
Dim volume, openingPrice, closingPrice, yearlyChange, perChange, negMaxPerChange, posMaxPerChange As Double
Dim tickerName,tickerNameForVolume,tickerNameForPerInc,tickerNameForPerDec As String

For Each ws in Worksheets
    'Initialize counter for the new table row and other variables.
    totalVolumeTableRow = 1
    volume = 0
    openingPrice = 0
    closingPrice = 0
    perChange = 0
    openingPriceChk = False
    ' Count the last row of the main table.
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"

' Iterate throgh the each row to find opening price, closing price, percentage change of each Ticker.
For i = 2 To lastRow
    ' Check that next row contains same Ticker or different. 
    ' If it is same then add the stock volume to get the total stock volume of that ticker.
    ' Here it won't add volume of last row.
    If ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value Then
        volume = volume + ws.Cells(i, 7).Value
        ' Check weather opening price is zero or not, to avoid divide by zero error at the time of calculating percentage change.
        ' If it is not zero and flag "openingPriceChk" = 0 then get the first non zero value. and set "openingPriceChk" to "1". so it won't overwrite the opening price value.
        If ws.Cells(i,3).Value <> 0 and openingPriceChk = True Then
        openingPrice = ws.Cells(i,3).Value
        openingPriceChk = True
        End If
    Else
    ' Next row Ticker is different than the current one. So get the ticker name, closing price, yearly change, percentage change and write to proper cells.
        ' Increament new table row by one because of header.
        totalVolumeTableRow = totalVolumeTableRow + 1
        ' Write current Ticker name, once the Ticker gets change in next row.
        tickerName = ws.Cells(i, 1).Value
        ws.Cells(totalVolumeTableRow, 9).Value = tickerName
        ' Write totalStockVolume for current Ticker after adding the last row volume.
        ws.Cells(totalVolumeTableRow, 12).Value = volume + ws.Cells(i, 7).Value
        ' find the closing price at the last row of each Ticker
        closingPrice = ws.Cells(i, 6).Value
        ' MsgBox ("Closing price: " & closingPrice & " Opening Price: " & openingPrice & " For Ticker: " & Cells(i, 1).Value)
        yearlyChange = closingPrice - openingPrice
        ws.Cells(totalVolumeTableRow, 10).Value = yearlyChange
        ' Check if yearlyChange is zero (means Opening price and closing price both are zero or both remains at the same price.) then perChange will be zero.
        If yearlyChange = 0 Then
        perChange = 0
        Else
        perChange = yearlyChange / openingPrice
        End If
        If i = 40000 Then
        MsgBox("Per change: " &perChange)
        End If
        ' Write the percentage change to proper cell and format tio percentage.
        ws.Cells(totalVolumeTableRow, 11).Value = perChange
        ws.Cells(totalVolumeTableRow, 11).NumberFormat = "0.00%"
        ' If percentage change is positive, show them with green color. If negative then show them with red color.
        if perChange>=0 Then
        ws.Cells(totalVolumeTableRow, 11).Interior.ColorIndex = 4
        Else
        ws.Cells(totalVolumeTableRow, 11).Interior.ColorIndex = 3
        End If
        openingPriceChk = False
        openingPrice = 0
        volume = 0
    End If
Next i

' Set the header for Summary table.
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % decrease"
ws.Range("O4").Value = "Greatest Total Volume"

' Find the Max value from total volume and get associated Ticker name.
' Find the Max value of Positive and negative percentage change and get associated Ticker name.
' maxTotalStockVolume = WorksheetFunction.Max(ws.Range("L2:L" & totalVolumeTableRow).Value)
maxTotalStockVolume = 0
posMaxPerChange = 0
negMaxPerChange = 0

For i = 2 To totalVolumeTableRow
If ws.Range("L" & i).Value > maxTotalStockVolume Then
    maxTotalStockVolume = ws.Range("L" & i).Value
    tickerNameForVolume = ws.Range("I" & i).Value
End If
If ws.Range("K" & i).Value > maxPerChange Then
    posMaxPerChange = ws.Range("K" & i).Value
    tickerNameForPerInc = ws.Range("I" & i).Value

ElseIf ws.Range("K" & i).Value < minPerChange Then
    negMaxPerChange = ws.Range("K" & i).Value
    tickerNameForPerDec = ws.Range("I" & i).Value
End If 
Next i
MsgBox("Iteration over new table is over.")
MsgBox("Ticker for total Volume: " & tickerNameForVolume)
MsgBox("Ticker for per Increase: " & tickerNameForPerInc)
MsgBox("Ticker for Per Decrease: " & tickerNameForPerDec)
    ws.Range("P4").Value = tickerNameForVolume
    ws.Range("P2").Value = tickerNameForPerInc
    ws.Range("P3").Value = tickerNameForPerDec
    ws.Range("Q4").Value = maxTotalStockVolume

    ws.Range("Q2").Value = posMaxPerChange
    ws.Range("Q3").Value = negMaxPerChange
    ws.Range("Q2:Q3").NumberFormat = "0.00%"
ws.Columns("A:Q").AutoFit
Next ws

End Sub
