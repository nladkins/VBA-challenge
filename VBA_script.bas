Attribute VB_Name = "Module1"
Option Explicit

Sub Multiple_year_stock_data()
        
        'Create initial variables needed to evaluate ticker data
        Dim i As Double
        Dim Summary_i As Double
        Dim LastRow As Double

        'Create variables needed to develop summary table
        Dim open_price As Double
        Dim close_price As Double
        Dim price_change As Double
        Dim percent_change As Double
        Dim volume As Double
        Dim percent_price_change As Double
        

        'Fill in summary column names
        Cells(1, 10).Value = "Ticker"
        Cells(1, 11).Value = "Yearly Change"
        Cells(1, 12).Value = "Percent Change"
        Cells(1, 13).Value = "Total Stock Change"

        ' Fill in the summary column names for the second summary table
        ' Cells(1, 15).Value = "Greatest % Increase"
        ' Cells(1, 16).Value = "Greatest % Decrease"
        ' Cells(1, 17).Value = "Greatest Total Volume"
        ' Cells(1, 18).Value = "Greatest Total Volume"

        ' Set the values for open price and total volume
        open_price = 0
        volume = 0
        Summary_i = 1

        ' Set variable for last row
        LastRow = Cells(Rows.Count, "A").End(xlUp).Row
        
        ' Start Loop through data table
        For i = 2 To LastRow
        
            ' capture the volume
            volume = volume + Cells(i, 7).Value

            ' Determine open price of the initial ticker
            If Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
            
                open_price = Cells(i, 3).Value
            
            End If
            
            ' Compare Ticker names in Column A
            If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then

                ' Enter Ticker names to the Summary Table in Column J
                Cells(Summary_i + 1, 10).Value = Cells(i, 1).Value

                ' Calculate closing price
                close_price = Cells(i, 6).Value

                ' Calculate Yearly Change by subtracting closing price from opening price
                price_change = close_price - open_price

                ' Enter Yearly Change to the Summary Table in Column K
                Cells(Summary_i + 1, 11).Value = price_change

                    ' Color the Yearly Cells based on their outcomes in Column K
                    If price_change < 0 Then
                    
                        Cells(Summary_i + 1, 11).Interior.ColorIndex = 3
                    
                    Else
                    
                        Cells(Summary_i + 1, 11).Interior.ColorIndex = 4

                    End If

                    ' Calculate the percent of annual change including a check for when open_price = 0
                    If open_price <> 0 Then
                        
                        percent_price_change = price_change / open_price

                        ' Enter % of change in Column L
                        Cells(Summary_i + 1, 12).NumberFormat = ".00%"
                        
                        ' insert percent price change
                        Cells(Summary_i + 1, 12).Value = percent_price_change

                    End If

                    ' Color the Yearly Cells based on their outcomes in Column K
                    If percent_price_change < 0 Then
                    
                        Cells(Summary_i + 1, 12).Interior.ColorIndex = 3
                    
                    Else
                    
                        Cells(Summary_i + 1, 12).Interior.ColorIndex = 4

                    End If
                    
                ' Calculate the Total Volume of the Ticker in Column G (MAY NEED TO UNDO)
                Cells(Summary_i + 1, 13).Value = volume
                
                volume = 0

                ' Enter total stock change in M (CHECK WHETHER THE DECIMAL SHOULD BE A COMMA)
                ' Cells(Summary_i + 1, 13).NumberFormat = "$0,00"

                Summary_i = Summary_i + 1
                
            End If
            
        Next i
        
End Sub

