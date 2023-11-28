Sub stocks_all_years()
    'run across all worksheets
    For Each ws In Worksheets
    
    'define columns
    ws.range("I1").Value = "Ticker"
    ws.range("J1").Value = "Yearly Change"
    ws.range("K1").Value = "Percent Change"
    ws.range("L1").Value = "Total Stock Volume"
    ws.range("P1").Value = "Ticker"
    ws.range("Q1").Value = "Value"
    ws.range("O2").Value = "Greatest % Increase"
    ws.range("O3").Value = "Greatest % Decrease"
    ws.range("O4").Value = "Greatest total Volume"
    
    'declare variables
    Dim Ticker_Symbol As String
    Dim close_price As Double
    Dim Stock_Total_Volume As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim Summary_Row As Integer
    Dim opening_price_index As Double
    Dim opening_price As Double
    Dim greatest_decrease As Double
    Dim greatest_increase As Double
    Dim greatest_volume As Double
    Dim greatest_decrease_ticker As String
    Dim greatest_increase_ticker As String
    Dim greatest_volume_ticker As String

    'set initial values of variables
    Summary_Row = 2
    Stock_Total_Volume = 0
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    yearly_change = 0
    percent_change = 0
    opening_price_index = 2
    
    'loop for each row starting at row 2 until last poulated row
    For i = 2 To lastrow
    
        'if current cell value is not equal to next cell value, perform this...
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
            Ticker_Symbol = ws.Cells(i, 1).Value
            close_price = ws.Cells(i, 6).Value
            opening_price = ws.Cells(opening_price_index, 3).Value
                
            Stock_Total_Volume = Stock_Total_Volume + ws.Cells(i, 7).Value
            yearly_change = close_price - opening_price
            percent_change = ((close_price - opening_price) / opening_price) * 100
            
            opening_price_index = i + 1
            
            'populating cells with values
            ws.range("I" & Summary_Row).Value = Ticker_Symbol
            ws.range("J" & Summary_Row).Value = yearly_change
            ws.range("K" & Summary_Row).Value = percent_change
            ws.range("L" & Summary_Row).Value = Stock_Total_Volume

            'moving to next row
            Summary_Row = Summary_Row + 1
            
            'resetting variables
            Stock_Total_Volume = 0
            yearly_change = 0
            percent_change = 0
            
        Else
           'add total volums cells
            Stock_Total_Volume = Stock_Total_Volume + ws.Cells(i, 7).Value
            
        End If
        
    Next i
    
    'find the largest value and the smallest value
    Set range_percent = ws.range("K:K")
    greatest_decrease = Application.WorksheetFunction.Min(range_percent)
    greatest_increase = Application.WorksheetFunction.Max(range_percent)
    
    'find largest value volume
    Set range_volume = ws.range("L:L")
    greatest_volume = Application.WorksheetFunction.Max(range_volume)
    
    'locate cell with smallest number and find ticker associated
    Set cell_decrease = ws.Cells.Find(greatest_decrease)
    greatest_decrease_ticker = cell_decrease.Offset(0, -2)

    'locate cell with largest number and find ticker associated
    Set cell_increase = ws.Cells.Find(greatest_increase)
    greatest_increase_ticker = cell_increase.Offset(0, -2)
     
     'locate cell with largest volume number and find ticker associated
    Set cell_volume = ws.Cells.Find(greatest_volume)
    greatest_volume_ticker = cell_volume.Offset(0, -3)
     
    'populate cells with values
    ws.range("Q3").Value = greatest_decrease
    ws.range("Q2").Value = greatest_increase
    ws.range("Q4").Value = greatest_volume
    
    ws.range("P2").Value = greatest_increase_ticker
    ws.range("P3").Value = greatest_decrease_ticker
    ws.range("P4").Value = greatest_volume_ticker
    
    'find last row in columns J and K
    j_lastrow = ws.Cells(Rows.Count, "J").End(xlUp).Row
    k_lastrow = ws.Cells(Rows.Count, "K").End(xlUp).Row
    
    'loop through column J and set conditional formatting
    For Each cell In ws.range("J2:J" & j_lastrow)
        If cell.Value < 0 Then
            cell.Interior.ColorIndex = 3
        ElseIf cell.Value > 0 Then
            cell.Interior.ColorIndex = 4
        Else
            cell.Interior.ColorIndex = 0
        End If
    Next cell
        
    'loop through column K and set conditional formatting
     For Each cell In ws.range("K2:K" & k_lastrow)
        If cell.Value < 0 Then
            cell.Interior.ColorIndex = 3
        ElseIf cell.Value > 0 Then
            cell.Interior.ColorIndex = 4
        Else
            cell.Interior.ColorIndex = 0
        End If
    Next cell
    
'adjust cell size to fit values
ws.Cells.EntireColumn.AutoFit
ws.Cells.EntireRow.AutoFit

Next ws
End Sub


