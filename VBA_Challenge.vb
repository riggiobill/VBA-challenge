

Sub Stocks()

Dim ws As Worksheet
Dim tickername As String
Dim NumberofStocks As Integer

Dim numDate As Long
Dim lowDate As Long
lowDate = 500000000                 'lowDate is impossibly high and highDate impossibly low so as to guarantee they always
Dim highDate As Long                'trigger their necessary conditions in the code. As such, the first date is always 
highDate = -1                       'going to be greater than highDate and less than lowDate.

Dim earlyOpeningPrice As Double
Dim lateClosingPrice As Double

Dim percentChange As Double
Dim totalVolume As Double





For i = 1 To 3 'change this second argument to whatever the max amount of sheets is

    Sheets(i).Activate
    
    'add four columns : Ticker, Yearly Change, Percent Change, and Total Stock Volume
    Sheets(i).Range("I1").Value = "Ticker"
    Sheets(i).Range("J1").Value = "Yearly Change"
    Sheets(i).Range("K1").Value = "Percent Change"
    Sheets(i).Range("L1").Value = "Total Stock Volume"
    NumberofStocks = 1
    totalVolume = 0         
    

                        'The loop below passes "n" as a variable to scan through each row up to a very high number. It is set
                        'to break the loop once it encounters a blank row and therefore smoothly accesses each row in the data
                        'set. It expects several sets of cases : finding continued entries for a stock ticker, finding a new 
                        'stock ticker entry, or encountering a blank row. When the program detects a new stock ticker, it 
                        'commits the data it's been storing in its variables numDate, lowDate, highDate, earlyOpeningPrice, and
                        'lateClosingPrice, resets the variables in preparation for incrementation, and smoothly continues through the
                        'loop. It then uses the variable NumberofStocks to control a reference point to how many stock tickers have been
                        'noted and summarized in the added columns. These checks and commits are listed below.


    For n = 2 To 1000000
        If Sheets(i).Cells(n, 1).Value = Sheets(i).Cells(n + 1, 1).Value Then  
            totalVolume = totalVolume + Sheets(i).Cells(n, 7).Value
            numDate = Sheets(i).Cells(n, 2).Value
            
            'highDate and lowDate are updated and checked to control the timeframe of each data entry
            If numDate < lowDate Then
                lowDate = numDate
                earlyOpeningPrice = Sheets(i).Cells(n, 3).Value
            End If
            
            If numDate > highDate Then
                highDate = numDate
                lateClosingPrice = Sheets(i).Cells(n, 6).Value
            End If
            
            
        'The trigger for this Elseif is : "If the current cell is different from the upcoming cell AND that cell isn't blank"
        ElseIf (Sheets(i).Cells(n, 1).Value <> Sheets(i).Cells(n + 1, 1).Value) And (Sheets(i).Cells(n + 1, 1).Value <> "") Then
            'With a new stock ticker detected, increment NumberofStocks and sum the values for totalVolume together
            NumberofStocks = NumberofStocks + 1
            totalVolume = totalVolume + Sheets(i).Cells(n, 7).Value
            'add, finalize value comparisons, reassign & reset variables, increment on as new passing to above
            ' modify Row Z, columns I, J, K, and L
            'assign price change values here before resetting variables to new data
            
            numDate = Sheets(i).Cells(n, 2).Value
            
            If numDate < lowDate Then
                lowDate = numDate
                earlyOpeningPrice = Sheets(i).Cells(n, 3).Value
            End If
            
            If numDate > highDate Then
                highDate = numDate
                lateClosingPrice = Sheets(i).Cells(n, 6).Value
            End If
            
            
            'assign difference in stock price, then CondForm for color based on change
                Sheets(i).Cells(NumberofStocks, 10).Value = lateClosingPrice - earlyOpeningPrice
                If (lateClosingPrice - earlyOpeningPrice < 0) Then
                    'color cell red
                    Sheets(i).Cells(NumberofStocks, 10).Interior.ColorIndex = 3
                Else
                    'color cell green
                    Sheets(i).Cells(NumberofStocks, 10).Interior.ColorIndex = 4
                End If
                
                
            'assign percentage gain or loss
            If (earlyOpeningPrice <> 0) Then
                    percentChange = (lateClosingPrice / earlyOpeningPrice) - 1
                    Sheets(i).Cells(NumberofStocks, 11).Value = percentChange
            Else
                Sheets(i).Cells(NumberofStocks, 11).Value = "n/a"
            End If
            
            'add total volume
            Sheets(i).Cells(NumberofStocks, 12).Value = totalVolume
            totalVolume = 0
            
            
            'reset date variables for next loop
            lowDate = 500000000
            highDate = -1
            
            tickername = Sheets(i).Cells(n, 1).Value
            Sheets(i).Cells(NumberofStocks, 9).Value = tickername
        
        Else
            NumberofStocks = NumberofStocks + 1
            'add and increment
            totalVolume = totalVolume + Sheets(i).Cells(n, 7).Value
            'assign price change values here before ending loop
            numDate = Sheets(i).Cells(n, 2).Value
            
            If numDate < lowDate Then
                lowDate = numDate
                earlyOpeningPrice = Sheets(i).Cells(n, 3).Value
            End If
            
            If numDate > highDate Then
                highDate = numDate
                lateClosingPrice = Sheets(i).Cells(n, 6).Value
            End If
            
            'assign difference in stock price, then CondForm for color based on change
                Sheets(i).Cells(NumberofStocks, 10).Value = lateClosingPrice - earlyOpeningPrice
                If (lateClosingPrice - earlyOpeningPrice < 0) Then
                    'color cell red
                    Sheets(i).Cells(NumberofStocks, 10).Interior.ColorIndex = 3
                Else
                    'color cell green
                    Sheets(i).Cells(NumberofStocks, 10).Interior.ColorIndex = 4
                End If
            
            'assign percentage gain or loss
            If (earlyOpeningPrice <> 0) Then
                    percentChange = (lateClosingPrice / earlyOpeningPrice) - 1
                    Sheets(i).Cells(NumberofStocks, 11).Value = percentChange
            Else
                Sheets(i).Cells(NumberofStocks, 11).Value = 0
            End If
            
            'add total volume
            Sheets(i).Cells(NumberofStocks, 12).Value = totalVolume
            totalVolume = 0
            
            tickername = Sheets(i).Cells(n, 1).Value
            Sheets(i).Cells(NumberofStocks, 9).Value = tickername
            lowDate = 500000000
            highDate = -1
            Exit For
            'add, finalize value comparisons, modify Row Z, columns I, J, K, and L, exit for
        End If

    Next n
    
    'applies formatting to percent change column
    Sheets(i).Columns(11).Style = "Percent"
    
Next i


End Sub
