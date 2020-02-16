

Sub VBAStocks_Multiple()

''------------------------------------------
'INSTRUCTIONS
'Create a script that will loop through all the stocks for one year for each run and take the following information.
    ' Ticker
    'Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
    'The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
    'The total stock volume of the stock.

''------------------------------------------
'Create a location for the summary table and its headers.
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"

'Create a location for the Greatest Values Table
Cells(2, 14).value = "Greatest % Increase"
Cells(3, 14).value = "Greatest % Decrease"
Cells(4, 14).value = "Greatest Total Volume"
 Range("N1:N4").ColumnWidth = 20

'Set a variable for multiple worksheets
Dim ws As Worksheet

'Keep track of the location for each ticker in the summary table.
Dim Summary_Table_Row As Double
Summary_Table_Row = 2

''------------------------------------------
'Assigning variables for holding and calculating Summary Table Values, including: Ticker, Yearly Change, Percent Change, and Total Stock Value
Dim Ticker As String
Dim Open_Date As Long
Dim Close_Date As Long
Dim Start_Price As Double
Dim End_Price As Double
Dim Change_Price As Double
Dim Percent_Change As Double
Dim TS_Volume As Double


''------------------------------------------
'Loop through all stocks
For Each ws In Worksheets
    
    Dim lastrow As Long
    lastrow = ws.Range("A1").End(xlDown).Row
    
    For I = 2 To lastrow

            'Check to see if we are within the same ticker symbol. If we are not...
            If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then

                'Set the ticker symbol name
                Ticker = Cells(I, 1).Value

                'Add to the total stock volume
                TS_Volume = TS_Volume + Cells(I, 7).Value

                'Store the end of year price
                End_Price = Cells(I, 6).Value

                'Print the ticker symbol in the Summary Table
                Range("I" & Summary_Table_Row).Value = Ticker

                'Calculate the Yearly Change
                Change_Price = End_Price - Start_Price

                'Calculate the Percent Change
                ''NOTE: Will not calculate due to 0 as starting value for Start_Price
                'Percent_Change = Change_Price / Start_Price

                'Print the yearly change in the Summary Table
                Range("J" & Summary_Table_Row).Value = Change_Price

                    If Range("J" & Summary_Table_Row).Value < 0 Then
                    Range("J" & Summary_Table_Row).Interior.Color = vbRed
                    Else: Range("J" & Summary_Table_Row).Interior.Color = vbGreen
                    End If

                'Print the percent change in the Summary Table
                Range("K" & Summary_Table_Row).Value = Percent_Change

                'Print the total stock volume in Summary Table
                Range("L" & Summary_Table_Row).Value = TS_Volume

                'Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1

                'Reset the ticker symbol
                Ticker = 0

                'Reset the Start Price
                Start_Price = Cells(I + 1, 3).Value

                'Reset the End Price
                End_Price = 0

                'Reset the Total Stock Volume
                TS_Volume = 0

                Else

                    'Add to the Total Stock Volume
                    TS_Volume = TS_Volume + Cells(I, 7).Value


            End If

    Next I

        'Inputs into Greatest Values Table
        'Greatest % Increase
        'Greatest_max = WorksheetFunction.Max(ws.cells(range("K"))
        'ws.cells(range("K")).value = Greatest_max

        'Greatest % Decrease
        'Greatest_min = WorksheetFunction.Min(cells(range("K")))
        'ws.cells(range("K")).value = Greatest_min

        'Greatest Total Volume
        'Dim Greatest_vol As Integer
        'Greatest_vol = WorksheetFunction.Max(ws.Range("L1").End(xlDown).Row)
       'ws.Range("L1").End(xlDown).Value = Greatest_vol
       'Print Greatest Total Volumen
       'Greatest_vol = Cells(4, 15).Value

Next ws

End Sub




