Sub Loop_WorkSheet():

Dim ws As Worksheet

    'Go to all worksheet
    For Each ws In Worksheets
        ws.Select
        'Run the formula
        Call Multiple_year_stock_data
    Next ws
    
End Sub


Sub Multiple_year_stock_data()

Dim rowcount As Double
rowcount = Cells(Rows.Count, 1).End(xlUp).Row

'### FORMATTING ###'
'Headers
[I1] = "Ticker"
[J1] = "Yearly Change"
[K1] = "Percent Change"
[L1] = "Total Stock Volume"
[O1] = "Ticker"
[P1] = "Value"
[N2] = "Greatest % Inc"
[N3] = "Greatest % Dec"
[N4] = "Greatest Total Vol"
'Headers I1 to  P1 Bold
Range("A1:P1").Font.Bold = True
'Autofit all columns and text center
Columns("A:P").AutoFit

'Copy <ticker> column to Ticker column
Range("I2").Value = Range("A2").Value
'Set row for next ticker to be copied to
tickerrow = 2


'### Declare Variables ###'
'Calculate yearlychange
Dim openamt, closeamt, yearlychange As Double
'Calculate percentchange
Dim percentchange As Double
'Calculate Total Stock
Dim totalstockvalue As LongLong

'Set initial value for openamt
openamt = [C2]
'Set initial value to zero
totalstockvalue = 0

'Loop through all tickers
For i = 2 To (rowcount)

    'Add row's value to stockvalue
    totalstockvalue = totalstockvalue + Cells(i, 7).Value
    
    'Check for new ticker to copy to table
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
    'Print totalstockvalue to table
    Cells(tickerrow, 12) = totalstockvalue
    
    'Reset totalstockvalue to 0
    totalstockvalue = 0
    
    'Create var and store the new ticker
    Dim ticker As String
    ticker = Cells(i, 1).Value
    
    'Print new ticker to table
    Range("I" & tickerrow).Value = ticker
    
    'Store closing price
    closeamt = Cells(i, 6).Value
    
    'Calculate yearly change and print to table
    yearlychange = closeamt - openamt
    Cells(tickerrow, 10).Value = yearlychange
    
    'Calculate percent change and print to table
    If openamt = 0 Then
    Cells(tickerrow, 11).Value = "NULL"
    
    Else
    percentchange = (closeamt - openamt) / openamt
    Cells(tickerrow, 11).Value = percentchange
    
    End If
    
    'Reset openamt price value for next ticker
    openamt = Cells(i + 1, 3).Value
    
     'Add one to tickerrow
    tickerrow = tickerrow + 1
    
    'If ticker is same as previous row
    Else
    
    End If
    
Next i

'Count number of rows in table
Dim formatrowcount As Integer
formatrowcount = Cells(Rows.Count, "I").End(xlUp).Row

'Format Column K as % with 2 digits
Range("K2:K" & formatrowcount).NumberFormat = "0.00%"

'Format new columns so text is centered
Range("I1:L" & formatrowcount, "N1:P4").HorizontalAlignment = xlCenter

'Add commas to volume numbers to make more readable
Range("L2:L" & formatrowcount).NumberFormat = "###,###,###,##0"

'Add conditional formatting to Column J, yearly change
For j = 2 To formatrowcount
    If Cells(j, 10).Value < 0 Then
    Cells(j, 10).Interior.ColorIndex = 3
    Else: Cells(j, 10).Interior.ColorIndex = 4
    End If
Next j

'Declare variables and create For loop to determine row with greatest % increase
Dim maxpercent As Double
Dim maxticker As String

maxpercent = 0.001

For r = 1 To formatrowcount
    If (Cells(r, 11).Value <> "NA") Then
        If (Cells(r, 11).Value > maxpercent) Then
            maxpercent = Cells(r + 1, 11).Value
            maxticker = Cells(r, 9).Value
        End If
    ElseIf (Cells(r, 11).Value = "NA") Then
    End If
Next r

'Print values to table for greatest % increase
[O2] = maxticker
[P2] = maxpercent

'Declare variables and create For loop to determine row with greatest % increase
Dim minpercent As Double
Dim minticker As String

minpercent = 0

For a = 2 To formatrowcount
    If (Cells(a, 11).Value < minpercent) Then
        minpercent = Cells(a, 11).Value
        minticker = Cells(a, 9).Value
    End If
Next a

'Return values to greatest % increase table
[O3] = minticker
[P3] = minpercent

'Define variables and loop to determine stock with greatest total volume
Dim maxvolume As LongLong
Dim maxvolticker As String

maxvolume = 1

For y = 2 To formatrowcount
    If (Cells(y, 12).Value > maxvolume) Then
        maxvolume = Cells(y, 12).Value
        maxvolticker = Cells(y, 9).Value
    End If
Next y

'Print highest total volume to table
[O4] = maxvolticker
[P4] = maxvolume

'Update formatting for highest Inc, Total Vol, and Dec
Range("P2:P3").NumberFormat = "0.00%"
Range("O2:P4").Font.Bold = False
Range("P4").NumberFormat = "###,###,###,#00"


End Sub
