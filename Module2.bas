Attribute VB_Name = "Module2"
Sub stock_summarizer()


    'Loop through all worksheets
    Dim ws_num As Integer
    Dim starting_ws As Worksheet
    Set starting_ws = ActiveSheet 'My macro buttons are in the first sheet and I want to go back to i=this sheet when the script is one
    ws_num = ThisWorkbook.Worksheets.Count
    
    Dim x As Integer
    
        For x = 1 To ws_num
            ThisWorkbook.Worksheets(x).Activate
            
            ' Set some variables for ticker symbol, ticker total, opening value, closing value, difference, Pchange, summary table row number
            Dim Ticker_Symbol As String
            Dim Ticker_Total As Double
            Dim Closing_Value As Double
            Dim Opening_Value As Double
            Dim Difference As Double
            Dim Pchange As Double
            Dim Summary_Table_Row As Integer
            
            'Set some initial values
            Ticker_Total = 0
            Summary_Table_Row = 2
            
            'Assign header values to Summary Table
            Cells(1, 10) = "Ticker"
            Cells(1, 11) = "Yearly Change"
            Cells(1, 12) = "Percent Change"
            Cells(1, 13) = "Total Stock Volume"

            
            ' Determine the Last Row for column A
            LastRow = Cells(Rows.Count, 1).End(xlUp).Row
            
            ' Loop through all rows
                For i = 2 To LastRow
                
                    ' Check if we are still within the same ticker symbol, if it is not ...
                    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                    
                        ' Set the Ticker symbol
                        Ticker_Symbol = Cells(i, 1).Value
                        Closing_Value = Cells(i, 6).Value
                        Difference = Closing_Value - Opening_Value
                    
                        'To deal with a #DIV/0 error, set Pchange to zero wherever Opening_Value is  = 0
                        If Opening_Value > 0 Then
                            Pchange = Difference / Opening_Value
                            Else
                            Pchange = 0
                        End If
                                
                        ' Print the Ticker Symbol, Difference, Pchange, and Ticker Total in the Summary Table
                        Range("J" & Summary_Table_Row).Value = Ticker_Symbol
                        Range("K" & Summary_Table_Row).Value = Difference
                        Range("L" & Summary_Table_Row).Value = Format(Pchange, "0.00%")
                        Range("M" & Summary_Table_Row).Value = Ticker_Total
                
                        ' Add one to the summary table row to go to the next row if ticker symbol changes
                        Summary_Table_Row = Summary_Table_Row + 1
                        ' Reset the Ticker Total
                        Ticker_Total = 0
                                        
                    Else
                
                        ' Add to the Ticker Total
                        Ticker_Total = Ticker_Total + Cells(i + 1, 7).Value
                
                    End If
                
                    ' To deal with the first row in the loop for opening value
                    If Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
                        Opening_Value = Cells(i, 3).Value
                    End If
                
                Next i
                    
                ' Determine the Last Row for the summary table (for purposes of the color formatting)
                LastRow2 = Cells(Rows.Count, 11).End(xlUp).Row

                'Just looping through to evaluate each value in this column for color designation
                For p = 2 To LastRow2
                    If Cells(p, 11).Value >= 0 Then
                        Cells(p, 11).Interior.ColorIndex = 4
                    Else
                        Cells(p, 11).Interior.ColorIndex = 3
                    End If
                Next p
            
            
        Next x
    
    'Activate the worksheet that was originally active (where our macro buttons are)
    starting_ws.Activate

End Sub




