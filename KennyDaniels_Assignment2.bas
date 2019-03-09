Attribute VB_Name = "Module1"
Sub main()
    Dim pointer As Double 'for figuring out what line im reading
    Dim Tickerpointer As Integer 'for figuring out what line im writing to
    Dim nextTicker As String 'for figuring out what symbol I'm on
    Dim tempVolume As Double 'for figuring out my total volume for a given symbol
    Dim lastRow As Double 'for holding the last row number
    Dim openAt As Double 'for holding what the stock opened at
    Dim closeAt As Double 'for holding what the stock closed at
    
    For Each ws In ActiveWorkbook.Worksheets 'applying code to all worksheets
        ws.Activate
        
        'setting up base cases
        Tickerpointer = 2
        pointer = 2
        nextTicker = Cells(pointer, 1)
        Cells(2, 9) = nextTicker
        openAt = Cells(2, 3).Value
    
        lastRow = Cells(Rows.Count, 1).End(xlUp).Row 'finding out what the last row number is
    
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
    

    
        While pointer < lastRow + 1
            'for if the symbol of the previous line and the current one match up, just update the volume
            If Cells(pointer, 1).Value = nextTicker Then
                tempVolume = tempVolume + Cells(pointer, 7).Value
                Cells(Tickerpointer, 12) = tempVolume
                closeAt = Cells(pointer, 6).Value
            'if the symbol of the previous line doesn't match the current, update all the variables
            ElseIf Cells(pointer, 1).Value <> nextTicker Then
                Cells(Tickerpointer, 10) = closeAt - openAt 'calculating different between opening vs closing price of the stock at the end of the year
                
                If (closeAt - openAt) > 0 Then 'coloring the cell based on gain or loss
                    Cells(Tickerpointer, 10).Interior.ColorIndex = 4
                ElseIf (closeAt - openAt) < 0 Then
                    Cells(Tickerpointer, 10).Interior.ColorIndex = 3
                End If
                
                If openAt <> 0 Then
                    Cells(Tickerpointer, 11) = (closeAt - openAt) / openAt 'calculating the percentage difference between the opening and closing prices
                Else
                    Cells(Tickerpointer, 11) = 0 'if the stock opens at 0, I need to use an else statement so that it catches a divide by 0 error
                End If
                tempVolume = Cells(pointer, 7).Value
                Tickerpointer = Tickerpointer + 1
            
                nextTicker = Cells(pointer, 1).Value
                Cells(Tickerpointer, 9) = nextTicker
            
                Cells(Tickerpointer, 12) = tempVolume 'updating new volume
                
                openAt = Cells(pointer, 3).Value 'assigning new opening price for stock
            End If
            If pointer = lastRow Then 'accounting for last case as the above loop ignores the last ticker for the else if
                Cells(Tickerpointer, 10) = closeAt - openAt
                
                If (closeAt - openAt) > 0 Then
                    Cells(Tickerpointer, 10).Interior.ColorIndex = 4
                ElseIf (closeAt - openAt) < 0 Then
                    Cells(Tickerpointer, 10).Interior.ColorIndex = 3
                End If
                    
                                If openAt <> 0 Then
                    Cells(Tickerpointer, 11) = (closeAt - openAt) / openAt
                Else
                    Cells(Tickerpointer, 11) = 0
                    End If
            End If
            pointer = pointer + 1
        Wend
        
        
        Dim bottomK As String
        bottomK = "K2:K" & Tickerpointer
        For Each cell In Range(bottomK)
            With cell
                .NumberFormat = "0.00%"
            End With
        Next
        
        Dim bottomJ As String
        bottomJ = "J2:J" & Tickerpointer
        For Each cell In Range(bottomJ)
            With cell
                .NumberFormat = "0.000000000"
            End With
        Next
    Next ws
    
    
    
End Sub
