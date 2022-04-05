Attribute VB_Name = "Module1"
Sub tickertotaler_moderate()



'define variable for calculations
Dim ws As Worksheet
Dim ticker As String
Dim vol As Double
Dim year_open As Double
Dim year_close As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim Summary_Table_Row As Integer

    ticker = " "
    vol = 0
    year_open = 0
    year_close = 0
    yearly_change = 0
    percent_change = 0


'run through each worksheet
For Each ws In ThisWorkbook.Worksheets
    'set headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"

    'setup the locations for variables
    Summary_Table_Row = 2

    'loop
        For I = 2 To ws.UsedRange.Rows.Count
             If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then
            
            'find the values
            ticker = ws.Cells(I, 1).Value
            vol = ws.Cells(I, 7).Value

            year_open = ws.Cells(I, 3).Value
            year_close = ws.Cells(I, 6).Value

            yearly_change = year_close - year_open
            percent_change = (yearly_change / year_open) * 100

            'insert values into summary
            ws.Cells(Summary_Table_Row, 9).Value = ticker
            ws.Cells(Summary_Table_Row, 10).Value = yearly_change
            ws.Cells(Summary_Table_Row, 11).Value = percent_change
            ws.Cells(Summary_Table_Row, 12).Value = vol
            Summary_Table_Row = Summary_Table_Row + 1

        
        End If


    'Color yearly change, Red Negative - Green Positive
    If (yearly_change > 0) Then
        ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
        
    ElseIf (yearly_change <= 0) Then
            ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
            
    End If
    

'finish loop
    Next I
    
ws.Columns("K").NumberFormat = "0.00%"


'move to next worksheet
Next ws


End Sub
