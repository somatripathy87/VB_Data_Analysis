

Sub CalculateStockSummary()
   

    'Defining all variables
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim arrData As Variant
    Dim i As Long, j As Long
    Dim ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim outputRow As Long
   
    'Variables to track the maximum and minimum values
    Dim maxPercentIncrease As Double
    Dim maxPercentIncreaseTicker As String
    Dim maxPercentDecrease As Double
    Dim maxPercentDecreaseTicker As String
    Dim maxTotalVolume As Double
    Dim maxTotalVolumeTicker As String

    'Apply loop to run through each sheet,
      For Each ws In Worksheets
        
    'Activate current sheet
        ws.Activate
       
    'Set reference to the active sheet
       Set ws = ActiveSheet

     ' Adding headers from columns I through L
    Range("I1:L1").Font.Bold = True
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"

    'Find the last row in column A
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    'Read all data in the sheet into an array
        arrData = ws.Range("A2:G" & lastRow).Value

    'Initialize output row to 2 to skip the header
        outputRow = 2
        
    'Initialize ticker value
        ticker = Range("A2").Value
        
    'Initialize open and close price
        openPrice = Range("C2").Value
        closePrice = Range("F2").Value

    'Initialize variables for maximum and minimum percent values and maximum volume
        maxPercentIncrease = 0
        maxPercentIncreaseTicker = ""
        maxPercentDecrease = 0
        maxPercentDecreaseTicker = ""
        maxTotalVolume = 0
        maxTotalVolumeTicker = ""


    'Loop through each row in the array
        For i = 1 To UBound(arrData)
    ' Skip if the row has been processed
            If arrData(i, 1) = "" Then GoTo ContinueLoop

    ' Get the values in the current row
            ticker = arrData(i, 1)
            openPrice = arrData(i, 3)
            closePrice = arrData(i, 6)
            
            ' Reset variables for each ticker
            
            percentChange = 0
            totalVolume = 0

            ' Loop through each row again to find the same ticker
            For j = 1 To UBound(arrData)
                If arrData(j, 1) = ticker Then
                    ' Calculate total volume
                    totalVolume = totalVolume + arrData(j, 7)
                    closePrice = arrData(j, 6)
                    ' Mark the ticker as processed
                    arrData(j, 1) = ""
                                       
                End If
            Next j

            'Calculate yearly change by subtracting last close price of the ticker with the first open price
            yearlyChange = closePrice - openPrice
            
            ' Calculate percent change and also look for avoiding division by zero
            percentChange = (yearlyChange / openPrice)
            

            ' Write results to columns I through L
            ws.Cells(outputRow, 9).Value = ticker
            ws.Cells(outputRow, 10).Value = yearlyChange
            ws.Cells(outputRow, 11).Value = percentChange
            ws.Cells(outputRow, 11).NumberFormat = "0.00%" 'Change format to show percentage
            ws.Cells(outputRow, 12).Value = totalVolume

            ' Color formatting
            If ws.Cells(outputRow, 10).Value > 0 Then
                ws.Cells(outputRow, 10).Interior.Color = RGB(0, 255, 0)
            Else
                ws.Cells(outputRow, 10).Interior.Color = RGB(255, 0, 0)
            End If

            ' Track the maximum and minimum values
            If percentChange > maxPercentIncrease Then
                maxPercentIncrease = percentChange
                maxPercentIncreaseTicker = ticker
            End If

            If percentChange < maxPercentDecrease Then
                maxPercentDecrease = percentChange
                maxPercentDecreaseTicker = ticker
            End If

            If totalVolume > maxTotalVolume Then
                maxTotalVolume = totalVolume
                maxTotalVolumeTicker = ticker
            End If

            ' Increment output row
            outputRow = outputRow + 1

ContinueLoop:
        Next i

        ' Print the greatest percentage increase, greatest percentage decrease, and greatest total volume
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        Range("P1:Q1").Font.Bold = True 'Change format of headers to show as Bold
        
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
               
        ws.Range("P2").Value = maxPercentIncreaseTicker
        ws.Range("P3").Value = maxPercentDecreaseTicker
        ws.Range("P4").Value = maxTotalVolumeTicker
        
        ws.Range("Q2").Value = maxPercentIncrease
        ws.Range("Q2").NumberFormat = "0.00%" 'Change format to show percentage
        ws.Range("Q3").Value = maxPercentDecrease
        ws.Range("Q3").NumberFormat = "0.00%" 'Change format to show percentage
        ws.Range("Q4").Value = maxTotalVolume
        
       Next ws
End Sub





