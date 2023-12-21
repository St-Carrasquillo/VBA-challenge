Attribute VB_Name = "Module1"
Sub YearlyChanges()

'creating loop for worksheets, must add 'ws.' before any cell or range reference
For Each ws In Worksheets

'declaring value titles
Dim Ticker As String

Dim YearlyChange As Double
YearlyChange = 0

Dim Percentage As Double
Percentage = 0
Dim lrgincrease As Double
Dim lrgdecrease As Double


Dim Total_Stock As LongLong
Total_Stock = 0
Dim hghstvlm As LongLong

Dim year_open As Double
year_open = 0

Dim year_close As Double
year_close = 0





' Keep track of the location for each ticker in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

'setting up last row lookup for loop
Dim lastRow As Long

lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'need to establish the first year open value
 year_open = ws.Cells(2, 3).Value
 
 
'Looping through all the rows and setting variable for year_open
    For i = 2 To lastRow
   
    'checking if ticker string has changed
    If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
    'setting variable values for adding to review table
    Ticker = ws.Cells(i, 1).Value
    
    Total_Stock = Total_Stock + ws.Cells(i, 7).Value
    
    year_close = ws.Cells(i, 6).Value
    
    YearlyChange = year_close - year_open
    
    Percentage = YearlyChange / year_open
  
    'adding data to summary table
    ws.Range("J" & Summary_Table_Row).Value = Ticker
    ws.Range("K" & Summary_Table_Row).Value = YearlyChange
        
        'formating yearlyChange cells to green if over 0 and red if under
        If YearlyChange > 0 Then
        ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
        Else: ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
        End If
    
    ws.Range("L" & Summary_Table_Row).Value = Format(Percentage, "0.00%")
    ws.Range("M" & Summary_Table_Row).Value = Total_Stock
        
    
    'resetting up position and values in summary table for next entry
    Summary_Table_Row = Summary_Table_Row + 1
    year_open = ws.Cells(i + 1, 3).Value
    Total_Stock = 0
    
    Else
    'if Ticker matches, continue adding to Total_Stock
    Total_Stock = Total_Stock + ws.Cells(i, 7).Value
    
    
    End If
Next i

ws.Range("J1").Value = "Ticker"
ws.Range("k1").Value = "Yearly Change"
ws.Range("l1").Value = "Percentage Change"
ws.Range("m1").Value = "Total Stock Volume"
ws.Range("P2").Value = "Greatest Increase"
ws.Range("P3").Value = "Greatest Decrease"
ws.Range("P4").Value = "Greatest Total Volume"
ws.Range("q1").Value = "Ticker"
ws.Range("r1").Value = "Value"


lrgincrease = WorksheetFunction.Max(ws.Range("L:L"))
lrgdecrease = WorksheetFunction.Min(ws.Range("L:L"))
ws.Range("r2").Value = Format(lrgincrease, "percent")
ws.Range("r3").Value = Format(lrgdecrease, "percent")
ws.Range("r4").Value = "=max(M:M)"
ws.Range("q2").Value = "=xlookup(R2,L:L,J:J)"
ws.Range("q3").Value = "=xlookup(R3,L:L,J:J)"
ws.Range("q4").Value = "=xlookup(R4,M:M,J:J)"
ws.Columns("J:R").AutoFit

Next ws

End Sub


