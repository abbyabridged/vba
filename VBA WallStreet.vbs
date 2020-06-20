Attribute VB_Name = "Module1"
Sub WallStreet()

'Loop through all worksheets using For Each method

Dim ws As Worksheet

For Each ws In Worksheets
ws.Activate
    
    'In each worksheet...

        'Set a variable for the ticker symbol
        Dim TickerSymbol As String
                  
        'Set a variable for total stock volume, starting from 0
        Dim TotalStockVolume As LongLong
        TotalStockVolume = 0
                                  
        'Track the row number for each ticker symbol in the summary table, starting from row 2
        Dim SummaryTableRow As Integer
        SummaryTableRow = 2
        
        'Set a variable for last row
        Dim lastrow As LongLong
        lastrow = Cells(Rows.Count, 1).End(xlUp).Row
        
        'Print header row for summary table (populate range with contents of an array)
        Dim SummaryHeaders As Variant
        SummaryHeaders = Array("Ticker", "Yearly Change", "Yearly Percentage Change", "Total Stock Volume")
        Range("I1:L1") = SummaryHeaders
        
    
        'Loop through each stock entry
        
        For I = 2 To lastrow
        
          'Compare ticker to previous entry. If not the same...
          If Cells(I - 1, 1).Value <> Cells(I, 1).Value Then
          
            'Set the opening price
            OpenPrice = Cells(I, 3).Value
          
          End If
            
            '--------------------------------
                                    
            'Compare ticker to next entry. If not the same...
            If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then
    
                'Set the ticker name
                TickerName = Cells(I, 1).Value
                                    
                'Add to the total stock volume
                TotalStockVolume = TotalStockVolume + Cells(I, 7).Value
            
                'Set the closing price
                ClosePrice = Cells(I, 6).Value
            
                'Set variable for yearly change
                Dim YearlyChange As Double
                YearlyChange = ClosePrice - OpenPrice
                
                
                'Set variable for percentage change
                Dim PercentChange As Double
                
                        'In instances where the opening and closing price are the same, Excel runs into an error as zero cannot  be divded.
                        'Use an if statement to set percentage change to 0 in these instances
                            If (ClosePrice - OpenPrice) = 0 Then
                            PercentageChange = 0
                            
                        'In instances where the opening price is zero, percentage change cannot be calculated
                        'Use an if statement to set percentage change to "N/A"
                            ElseIf OpenPrice = 0 Then
                            PercentageChange = "N/A"
                        
                        'For all other instances calculate percentage change
                            Else
                            PercentChange = ((ClosePrice - OpenPrice) / OpenPrice)
                            End If
                                    
                 'Print all the information required...
                     
                 'Print the ticker name in the summary table
                 Range("I" & SummaryTableRow).Value = TickerName
                 
                 
                 'Print the yearly change (opening price - closing price)
                 'Format to two decimal places
                 Range("J" & SummaryTableRow).Value = Format(YearlyChange, "#,##0.00")
                              
                 'Print the percentage change (closing price - opening price)/opening price
                 'Format as a percentage with two decimal places
                 Range("K" & SummaryTableRow).Value = Format(PercentChange, "Percent")
                   
                 'Print the total stock volume in the summary table
                 Range("L" & SummaryTableRow).Value = TotalStockVolume
             
                            
                 'Add to summary row and reset total stock volume...
                            
                 'Add 1 to the summary table row
                 SummaryTableRow = SummaryTableRow + 1
             
                 'Reset the total stock volume to zero
                 TotalStockVolume = 0
    
            End If
                
                '-------------------------------------
                
                ' Compare ticker to next entry. If the same...
                If Cells(I + 1, 1).Value = Cells(I, 1).Value Then
                        
                'Add to the total stock volume
                TotalStockVolume = TotalStockVolume + Cells(I, 7).Value
                                                                        
            End If
            
            'Apply conditional formatting for yearly change (red for negative, green for positive)
            
            If YearlyChange < 0 Then
            Range("J" & SummaryTableRow - 1).Interior.ColorIndex = 3
            
            ElseIf YearlyChange > 0 Then
            Range("J" & SummaryTableRow - 1).Interior.ColorIndex = 4
                         
            End If
                
                                                   
        Next I
            
'Create statistics table...

'Print header row for statistics table
Dim StatsHeaders As Variant
StatsHeaders = Array("Ticker", "Value")
Range("P1:Q1") = StatsHeaders

'Print header column for statistics table
Dim StatsLabels As Variant
StatsLabels = Array("Greatest % Increase", "Greatest % Decrease", "Greatest Total Volume")
StatsLabels = WorksheetFunction.Transpose(StatsLabels)
Range("O2:O4") = StatsLabels

'Use VBA Max function to find stock with greatest percentage increase and corresponding ticker symbol
'Print ticker symbol in cell P2
'Print percentage increase in Q2 formatted as percent

MaxPercentInc = WorksheetFunction.Max(Range("K:K"))
Range("Q2").Value = Format(MaxPercentInc, "Percent")

MaxPercentIncTicker = WorksheetFunction.Index(Range("I:I"), WorksheetFunction.Match(Cells(2, 17).Value, Range("K:K"), 0))
Range("P2").Value = MaxPercentIncTicker


'Use VBA Min function to find stock with greatest percentage decrease and corresponding ticker symbol
'Print ticker symbol in cell P3
'Print percentage decrease in Q3 formatted as a percentage

MaxPercentDec = WorksheetFunction.Min(Range("K:K"))
Range("Q3").Value = Format(MaxPercentDec, "Percent")

MaxPercentDecTicker = WorksheetFunction.Index(Range("I:I"), WorksheetFunction.Match(Cells(3, 17).Value, Range("K:K"), 0))
Range("P3").Value = MaxPercentDecTicker

            '---------
            
            'It occurred to me that the lowest value in this column could be positive, which means there is no percentage decrease.
            'If value found by Min function is positive, then there is no percentage decrease, so print "Not available"
                If MaxPercentDec >= 0 Then
                Range("Q3").Value = "No decreases"
                Range("P3").Value = "No decreases"
               
                End If
                
            '------------
    
    
'Use VBA Max function to find stock with greatest total volume and corresponding ticker symbol
'Print ticker symbol in cell P4
'Print total volume in Q4

MaxVolumeChange = WorksheetFunction.Max(Range("L:L"))
Range("Q4").Value = MaxVolumeChange

MaxVolTicker = WorksheetFunction.Index(Range("I:I"), WorksheetFunction.Match(Cells(4, 17).Value, Range("L:L"), 0))
Range("P4").Value = MaxVolTicker

'Autofit columns
Columns("I:Q").EntireColumn.AutoFit

Next ws
    

End Sub

