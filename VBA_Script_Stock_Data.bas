Attribute VB_Name = "Module1"
Sub Stock_Data():

    'Declaration of variables
    Dim i, j, tickerCount As Integer
    Dim open_price, close_price, yearly_change, percent_change As Double
    Dim max, min As Double
    Dim max_ticker, min_ticker, max_total_ticker As String
    Dim total, max_total As LongLong
    Dim ticker As String

    'assigning header values
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    
    'assigning initial values
    open_price = Cells(2, 3).Value
    printRowCount = 2
    total = 0
    
    'looping for storing all data
    For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row
               
        'storing values for total
        total = total + Cells(i, 7).Value
            
         If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            'print ticker column
            ticker = Cells(i, 1).Value
            Cells(printRowCount, 9).Value = ticker
                    
            'print yearly change column
            close_price = Cells(i, 6).Value
            yearly_change = close_price - open_price
            Cells(printRowCount, 10).Value = Round(yearly_change, 2)
                    
            'fill colors according to values
            If yearly_change < 0 Then
                Cells(printRowCount, 10).Interior.ColorIndex = 3
            Else
                Cells(printRowCount, 10).Interior.ColorIndex = 4
            End If
                    
            'make a column for percent change
            percent_change = yearly_change / open_price
            Cells(printRowCount, 11).Value = FormatPercent(percent_change)
                    
            'print total of the yearly volume
            Cells(printRowCount, 12).Value = total
            'reset the total volume
            total = 0
                  
            'assigning opening price for another year
            open_price = Cells(i + 1, 3).Value
            printRowCount = printRowCount + 1
        End If
               
    Next i
    
    'printing titles of the table
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    
    'assuming first row as maximun values
    max = Cells(2, 11).Value
    min = Cells(2, 11).Value
    max_total = Cells(2, 12).Value
    
    'condition to get values
    For i = 2 To Cells(Rows.Count, 9).End(xlUp).Row
        
        If Cells(i, 11).Value >= max Then
            max = Cells(i, 11).Value
            max_ticker = Cells(i, 9).Value
        ElseIf Cells(i, 11).Value <= min Then
            min = Cells(i, 11).Value
            min_ticker = Cells(i, 9).Value
        ElseIf Cells(i, 12).Value >= max_total Then
            max_total = Cells(i, 12).Value
            max_total_ticker = Cells(i, 9).Value
        End If
    
    Next i
    
    'final result
    Cells(2, 17).Value = FormatPercent(max)
    Cells(2, 16).Value = max_ticker
    Cells(3, 17).Value = FormatPercent(min)
    Cells(3, 16).Value = min_ticker
    Cells(4, 17).Value = max_total
    Cells(4, 16).Value = max_total_ticker
    
End Sub



