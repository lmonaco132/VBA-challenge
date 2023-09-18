Attribute VB_Name = "Module1"
Sub vba_challenge()

'set up variable for the for-each loop to loop through the sheets
Dim ws As Worksheet

'integers to hold two row indeces, last for two different columns
Dim last_row As LongLong
Dim new_last_row As LongLong

'loop counters (two different for loops)
Dim i As LongLong
Dim j As LongLong

'counter to know what row to write to for different ticker symbols
Dim row_counter As LongLong
'start at 2, first row holds column names
row_counter = 2

'vars to hold values within a block of same ticker symbol rows
Dim open_price As Double
Dim running_total As LongLong


'loop through the sheets in this workbook
For Each ws In ThisWorkbook.Worksheets
    
    'find the last row of column A (last row of existing data) in this sheet
    last_row = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    'set up the column labels for this sheet where i will write summary data
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Value"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    
    'set up first opening price
    open_price = ws.Cells(2, 3).Value
    
    'loop through the rows, starting with first value in 2 and going to last row
    For i = 2 To last_row
    
        'check if current ticker symbol is different from the next
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            'this is the case where the current and next ticker symbols are DIFFERENT
            
            'write ticker symbol to current row of summary column (row_counter)
            ws.Cells(row_counter, 9).Value = ws.Cells(i, 1).Value
            
            'add this last row's total volume to running total of volume, then write running total of vol to row_counter
            running_total = running_total + ws.Cells(i, 7).Value
            ws.Cells(row_counter, 12).Value = running_total
            'then reset running total of vol to 0 for next ticker symbol
            running_total = 0
            
            'calculate yearly difference using opening price that's stored and closing price from this row
            'write the yearly difference to row_counter column J (10)
            ws.Cells(row_counter, 10).Value = ws.Cells(i, 6).Value - open_price
            
            'calculate the percent change using yearly difference
            'percent change = new value - original value / original value
            ws.Cells(row_counter, 11).Value = ws.Cells(row_counter, 10).Value / open_price
            
            'set the formatting of percent change as a percent
            ws.Cells(row_counter, 11).NumberFormat = "0.00%"
            
            'then set opening price to next row (i+1) to set up for next set of rows with new ticker symbol
            open_price = ws.Cells(i + 1, 3).Value
            
            'set conditional formatting for yearly difference cell that was just written to based on what was written
            If ws.Cells(row_counter, 10).Value > 0 Then
                'this is the case where yearly change is positive, set fill color to green
                ws.Cells(row_counter, 10).Interior.ColorIndex = 4
            Else
                'this is the case where yearly change is negative (or exactly 0) set fill color to red
                ws.Cells(row_counter, 10).Interior.ColorIndex = 3
            End If
            
            
            'LAST thing to do in this if block is to UPDATE the row counter
            'so i don't overwrite these values when i write for the next ticker symbol
            'do this at the end because it's after i'm done accessing this row
            row_counter = row_counter + 1
            
        Else
            'this is the case where current and next ticker symbols are the SAME
            'i want to add the total volume from this row to the running total
            running_total = running_total + ws.Cells(i, 7).Value
            
        End If
        
        
    'end of for loop block, this will progress to next row
    Next i
    
    
    'i'm done with row_counter now, i need to reset it for the next sheet
    row_counter = 2
    
    'this is where i will find the max and min percent diff and max total volume
    'loop through the rows i just populated with data from the above for loop
    
    'to loop through those rows i need to define last_new_row (aka the last row of the columns i filled)
    new_last_row = ws.Cells(ws.Rows.Count, 9).End(xlUp).Row
    
    'to find the extremes, we set these values at first to the first values
    'percent change is set to same first value for both greatest inc and greatest dec
    'ticker is just set to first ticker, it will update when we find larger/smaller values to go there
    ws.Cells(2, 16).Value = ws.Cells(2, 1).Value
    ws.Cells(3, 16).Value = ws.Cells(2, 1).Value
    ws.Cells(4, 16).Value = ws.Cells(2, 1).Value
    ws.Range("Q2:Q3").Value = ws.Range("K2").Value
    ws.Range("Q4").Value = ws.Range("L2").Value
    'then we'll test each row and update if we find a value more extreme
    
    'use other loop counter for this
    For j = 2 To new_last_row
    
        'this is looping through the rows i wrote to, looking for max and min % change and max total vol
        
        'first check if percent change is higher than what's recorded
        If ws.Cells(j, 11).Value > ws.Range("Q2").Value Then
            'this is the case where the current row has a greater % inc than what's recorded
            'so we update the greatest % inc to be this value (with ticker symbol)
            ws.Cells(2, 16).Value = ws.Cells(j, 9).Value
            ws.Range("Q2").Value = ws.Cells(j, 11).Value
        ElseIf ws.Cells(j, 11).Value < ws.Range("Q3").Value Then
            'this is the case where the current row has lower % inc (aka higher % dec) than what's recorded
            'so we update the greatest % dec to be this value (with ticker symbol)
            ws.Cells(3, 16).Value = ws.Cells(j, 9).Value
            ws.Range("Q3").Value = ws.Cells(j, 11).Value
        End If
        
        
        'separate check, for if the current total volume is higher than recorded greatest
        If ws.Cells(j, 12).Value > ws.Range("Q4").Value Then
            'this is the case where the current total is greater than the recorded greatest total
            'so update the greatest total to be this value (with ticker symbol)
            ws.Cells(4, 17).Value = ws.Cells(j, 12).Value
            ws.Cells(4, 16).Value = ws.Cells(j, 9).Value
        End If
        
    'end of loop block, this will progress to the next row
    Next j
    
    'after having found the max and min % change, update those cells to percent format
    ws.Range("Q2:Q3").NumberFormat = "0.00%"
    
    
'this is the end of the for each loop block, progress to the next sheet
Next ws


End Sub

