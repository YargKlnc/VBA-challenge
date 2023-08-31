Attribute VB_Name = "Module1"
Sub Multiple_VBA_Challenge_by_YK_Still_Working()

    Dim Newsheet As Worksheet
    For Each Newsheet In ThisWorkbook.Worksheets
        Newsheet.Select
        
    'The script loops through one year of all stock data and reads/stores (outputs) all of the following values from each row:
        Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("2018")
        'ticker symbol (set a variable to hold ticker name)(how can I go from one ticker to other, detect when ticker changes)
        Dim Ticker_Name As String
        'set an initial variable for holding the total stock volume
        Dim Total_Stock_Volume As Double
        Total_Stock_Volume = 0
        'keep track of the location for each ticker name brand in the summary table
        Dim Summary_Table_Row As Long
        Summary_Table_Row = 2
 
        'set a variable for Opening Value
        Dim Opening_Value As Double
        'set a variable for Closing Value
        Dim Closing_Value As Double
        'set a variable for Percentage Change
        
        Dim Percent_Change As Double
        Dim Max_Change As Double
        Dim Min_Change As Double
        Dim Greatest_Total_Volume As Double
        
        Dim Greatest_Total_Volume_Ticker As Integer
        Dim Max_Change_Ticker As Integer
        Dim Min_Change_Ticker As Integer
        
        Dim Increase_number As Integer
                
        'Dim Greatest_Total_Volume_Ticker As String
        Dim Max_Cell As Range
        Dim Min_Cell As Range
        
        'define the last row
        Dim LastRow As Long
        Cells(Rows.Count, 1).End(xlUp).Select
        LastRow = Selection.Row
        'LastRow = Cells(Rows.Count, 1).End(xlUp)
        'LastRow = Cells(Rows.Count, A).End(xlUp).Row
        
        'define Opening_Value and Closing_Value (closing value is last closing value of the year)
        Opening_Value = Range("C2").Value
        'Closing_Value = Cells(2, 6).Value?
                          
        'Define Column K (Percent Change)
        Range("K:K").Select
        Selection.NumberFormat = "0.00%"
                        
        'give name to new cells
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        
        'loop through rows in column for change
        For i = 2 To LastRow
        
        Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
        
            'check if we are still within the same ticker name, if it is not change name
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then ' in final row for ticker
                ' get closing value
                Closing_Value = Cells(i, 6).Value
                ' calculate yearly change
                Yearly_Change = Closing_Value - Opening_Value
                ' put yearly change into summary table
                Range("J" & Summary_Table_Row).Value = Yearly_Change
                ' if to handle zero opening value
                Percent_Change = (Yearly_Change / Opening_Value)
                
                Range("K" & Summary_Table_Row).Value = Percent_Change
                
                'Max_Change = Application.WorksheetFunction.Max(Range("K:K"))
                'Cells(2, 18) = Max_Change
                
                'Min_Change = Application.WorksheetFunction.Min(Range("K:K"))
                'Cells(3, 18) = Min_Change
                
                'Greatest_Total_Volume = Application.WorksheetFunction.Max(Range("L:L"))
                'Cells(4, 18) = Greatest_Total_Volume
                
                If Cells(Summary_Table_Row, 10).Value < 0 Then
                    Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                    
                Else:
                    Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                End If
                
                'If Percent_Change > Max_Change Then
                    'Cells(2, 18).Value = Max_Change
                    'End If
                'buraya if yaparak max ve min change belirle
                    
                
                ' change opening value to be opening for next ticker
                Opening_Value = Cells(i + 1, 3).Value
                
                'set the ticker name
                Ticker_Name = Cells(i, 1).Value
        
                'add to the Total Stock Volume
                'Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
        
                'print the ticker name in the summary table
                Range("I" & Summary_Table_Row).Value = Ticker_Name
        
                'print the ticker name amount to the summary table
                Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
                
               'reset the Total Stock Volume
                Total_Stock_Volume = 0
                
                'add one to the summary table row, increment summary row
                Summary_Table_Row = Summary_Table_Row + 1

            End If
        Next i
        
        'give name to new cells
        Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(4, 15).Value = "Greatest Total Volume"
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"
        'Cells(2, 17).Value = "Max_Change"
        'Cells(3, 17).Value = "Min_Change"
        'Cells(4, 17).Value = "Greatest_Total_Volume"
                
                Range("Q2:Q3").Select
                Selection.NumberFormat = "0.00%"
                

                Max_Change = Application.WorksheetFunction.Max(Range("K:K"))
                Cells(2, 17) = Max_Change
                'Max_Change_Ticker = Cells(Summary_Table_Row, 1).Value
                'Cells(2, 16) = Max_Change_Ticker
                

                Min_Change = Application.WorksheetFunction.Min(Range("K:K"))
                Cells(3, 17) = Min_Change
                'Min_Change_Ticker = Cells(Summary_Table_Row, 1).Value
                'Cells(3, 16) = Min_Change_Ticker
                Increase_number = WorksheetFunction.Match(WorksheetFunction.Max(Range("K2:K" & LastRow)), Range("K2:K" & LastRow), 0)
                Range("P2") = Cells(Increase_number + 1, 9)
                
                Decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(Range("K2:K" & LastRow)), Range("K2:K" & LastRow), 0)
                Range("P3") = Cells(Decrease_number + 1, 9)
                
                Max_number = WorksheetFunction.Match(WorksheetFunction.Max(Range("L2:L" & LastRow)), Range("L2:L" & LastRow), 0)
                Range("P4") = Cells(Max_number + 1, 9)
                Greatest_Total_Volume = Application.WorksheetFunction.Max(Range("L:L"))
                
                
                'Greatest_Total_Volume_Ticker = Cells(Summary_Table_Row, 1).Value
                Cells(4, 17) = Greatest_Total_Volume
                'Cells(4, 16) = Greatest_Total_Volume_Ticker


    Next Newsheet

End Sub



