Sub StockData()

Dim openingPriceChk As Boolean
Dim totalVolumeTableRow, lastRow, openingPriceRow, maxTotalStockVolume, maxStockCell As Long
Dim volume, openingPrice, closingPrice, yearlyChange, perChange, negMaxPerChange, posMaxPerChange As Double
Dim tickerName,tickerNameForVolume,tickerNameForPerInc,tickerNameForPerDec As String
Dim StockDataArr As Variant
Dim SourceRange As Range

For Each ws in Worksheets
    'Initialize the variables and the counter for the new table row.
    totalVolumeTableRow = 1
    volume = 0
    openingPrice = 0
    closingPrice = 0
    perChange = 0
    openingPriceChk = False
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ' Count the last row of the main table.
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    ' copy the source data in array 
    Set SourceRange = ws.Range("A1:G" &lastRow)
    StockDataArr = SourceRange.value
' Iterate throgh the row of array to find opening price, closing price, percentage change of each Ticker.
For i = 2 To UBound(StockDataArr, 1)-1
    ' Check that next row contains same Ticker or different. 
    ' If it is same then add the stock volume to get the total stock volume of that ticker.
    ' Here it won't add volume of last row.
    If StockDataArr(i,1) = StockDataArr(i+1,1)  Then
        volume = volume + StockDataArr(i, 7)
        ' Check weather opening price is zero or not, to avoid divide by zero error at the time of calculating percentage change.
        ' If it is not zero and flag "openingPriceChk" = false then get the first non zero value. and set "openingPriceChk" to "1". so it won't overwrite the opening price value in next iteration.
        If ws.Cells(i,3).Value <> 0 and openingPriceChk = False Then
        openingPrice = StockDataArr(i, 3)
        openingPriceChk = True
        End If
    Else
        ' Here Next row Ticker is different than the current one. So get the current row ticker name, closing price, yearly change, percentage change and write to proper cells.
        ' Increament new table row by one because of header.
        totalVolumeTableRow = totalVolumeTableRow + 1
        ' Once the Ticker gets change in next row, Write current Ticker name.
        tickerName = StockDataArr(i, 1)
        ws.Cells(totalVolumeTableRow, 9).Value = tickerName
        ' Write totalStockVolume for current Ticker after adding the last row volume.
        ws.Cells(totalVolumeTableRow, 12).Value = volume + StockDataArr(i, 7)
        ' find the closing price at the last row of each Ticker
        closingPrice = StockDataArr(i, 6)
        yearlyChange = closingPrice - openingPrice
        ws.Cells(totalVolumeTableRow, 10).Value = yearlyChange
        ' Check if yearlyChange is zero (means Opening price and closing price both are zero or both remains at the same price.) then perChange will be zero.
        If yearlyChange = 0 Then
        perChange = 0
        Else
        perChange = yearlyChange / openingPrice
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
        ' Reset the flag, opening price and volume.
        openingPriceChk = False
        openingPrice = 0
        volume = 0
    End If
Next i

' Set the headers for Summary table.
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % decrease"
ws.Range("O4").Value = "Greatest Total Volume"

' Find the Max value from total volume column and get associated Ticker name.
' Find the Max value of Positive and negative percentage change and get associated Ticker name.

maxTotalStockVolume = 0
posMaxPerChange = 0
negMaxPerChange = 0

For i = 2 To totalVolumeTableRow
' If current volume is greater than "maxTotalStockVolume" then make it as "maxTotalStockVolume".
If ws.Range("L" & i).Value > maxTotalStockVolume Then
    maxTotalStockVolume = ws.Range("L" & i).Value
    tickerNameForVolume = ws.Range("I" & i).Value
End If
' If current volume is greater than "posMaxPerChange" then make it as "posMaxPerChange".
If ws.Range("K" & i).Value > posMaxPerChange Then
    posMaxPerChange = ws.Range("K" & i).Value
    tickerNameForPerInc = ws.Range("I" & i).Value
' If current volume is less than "negMaxPerChange" then make it as "negMaxPerChange".
ElseIf ws.Range("K" & i).Value < negMaxPerChange Then
    negMaxPerChange = ws.Range("K" & i).Value
    tickerNameForPerDec = ws.Range("I" & i).Value
End If 
Next i

    ws.Range("P4").Value = tickerNameForVolume
    ws.Range("P2").Value = tickerNameForPerInc
    ws.Range("P3").Value = tickerNameForPerDec
    ws.Range("Q4").Value = maxTotalStockVolume

    ws.Range("Q2").Value = posMaxPerChange
    ws.Range("Q3").Value = negMaxPerChange
    ws.Range("Q2:Q3").NumberFormat = "0.00%"
ws.Columns("A:Q").AutoFit
Erase StockDataArr
Next ws
End Sub
