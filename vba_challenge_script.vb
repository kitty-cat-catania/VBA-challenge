Sub plz()


'Loop through each worksheet
For Each ws In Worksheets
    ws.Range("K1") = "Ticker"
    ws.Range("L1") = "Yearly Change"
    ws.Range("M1") = "Percent Change"
    ws.Range("N1") = "Total Stock Volume"
    Dim RowCount As Long
    Dim i As Long
    'Unique is a variable for the  number of rows of unique tickers per ws
    Dim Unique As Integer
    
    Dim j As Integer
    Dim TckrCount As Integer
    Dim Tckr As String
    
    Dim calcTarget As String
    Dim stkVol As Double
     j = 2
    Dim k As Long
    
    Dim percentChng As Double
    
    Dim yrChng As Double
    
    Dim startVal As Double
    
    Dim endVal As Double
    
    Dim startD1 As Long
    Dim starD2 As Long
    Dim startD3 As Long
    
    Dim endD1 As Long
    Dim endD2 As Long
    Dim endD3 As Long
    
    'Sets the start date equal to the first date
    startD1 = 20160101
    
    startD2 = 20150101
    
    startD3 = 20140101
    'Sets the end date equal to the last date
    endD1 = 20161230
    
    endD2 = 20151231
    
    endD3 = 20141231
    
 
    'Set the stock volume equal to zero initially
    stkVol = 0
    
    'Sets number of rows of unique tickers equal to 2 initially
    Unique = 2
    
    'Store the number of rows in each ws in the variable RowCount
    RowCount = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Loop through all rows in the worksheet
    For i = 2 To RowCount
    
        'Check to see if we are still within the same ticker--if we are not then...
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
            stkVol = 0
           
            'Set the ticker name
            Tckr = ws.Cells(i, 1).Value
            
            'Add to the total stock volume
            stkVol = (stkVol + ws.Cells(i, 7).Value)
            
            
            
            'Print the volume to the stock volume column
            ws.Range("N" & Unique).Value = stkVol
            
            
            If ws.Cells(i, 2).Value = endD1 Then
            endVal = ws.Cells(i, 6).Value
            
            End If
            
            yrChng = startVal - endVal
            ws.Cells(Unique, 12).Value = yrChng
            percentChng = ((startVal - endVal) / startVal) * 100
            ws.Cells(Unique, 13).Value = percentChng
            
            
            
            Unique = Unique + 1
            
        
        Else
            'Add to the stock volume
            stkVol = stkVol + ws.Cells(i, 7).Value
            'Print the brand name to the Ticker column
            ws.Range("K" & Unique).Value = Tckr
            
            If ws.Cells(i, 2).Value = startD1 Then
                startVal = ws.Cells(i, 3).Value

            End If
            

        End If
    Next i
    TckrRow = ws.Cells(Rows.Count, 11).End(xlUp).Row
    

Next ws

End Sub


