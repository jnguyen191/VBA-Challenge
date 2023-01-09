# VBA-Challenge
Sub stock_analysis():


'STEP 1: Declare all variables
    'Use Long bc of overflow
    Dim lastRow As Long
    
    'totalVolume has large number, cannot use Integer. LongLong does not work (error: User-defined type not found)
    Dim totalVolume As String
    
    'openPrice has decimals
    Dim openPrice As Double
    
    'closePrice has decimals
    Dim closePrice As Double
    
    'use String because ticker includes text
    Dim Ticker As String
    
    'use Double because percent Change has decimals
    Dim percentChange As Double
    
    Dim summaryRow As Integer
    
    Dim Ticker2 As String
    Dim Value As Double
    

'to run loop through ALL worksheets
    For Each ws In Worksheets
    ws.Activate
    
'STEP 2: Set Initial values before running ForLoop, if neccessary
    
    'label summaryRow starting on J2
    summaryRow = 2
    
    'to count the last Row of the data
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    'set Cell(2,3) as opening price
    openPrice = Cells(2, 3).Value
    
    'beginning stock is 0
    totalVolume = 0
    
    Value = 0
    
    
'STEP 3: Set Column Headers for summary Table - set the title for columns to put the below If calculation results in
    Range("J1").Value = "Ticker"
    Range("K1").Value = "Yearly Change"
    Range("L1").Value = "Percent Change"
    Range("M1").Value = "Total Volume"
    
    'when using Range("L2:L" & summaryTable), it only format the Row 2 on ws B-F???
    Range("L2:L3500").NumberFormat = "0.00%"
    
    Range("Q1").Value = "Ticker2"
    Range("R1").Value = "Value"
    Range("P2").Value = "Greatest % increase"
    Range("P3").Value = "Greatest % decrease"
    Range("P4").Value = "Greatest Total Volume"
    
    
    
'STEP 3: Summary Table Calculations - loop through every row with data
    For currentRow = 2 To lastRow
    
        'Calculate totalVolume
        totalVolume = totalVolume + Cells(currentRow, 7)
        
    
        'Check if next Row and current Row have different Tickers: If the nextRow is not equal to the current Row, grab the ticker value (AAB, AAC, etc.); the closePrice of that row; the dollarsChange value
        If Cells(currentRow + 1, 1).Value <> Cells(currentRow, 1).Value Then
            'set the values for current ticker
            Ticker = Cells(currentRow, 1).Value
            closePrice = Cells(currentRow, 6).Value
            dollarsChange = closePrice - openPrice
            percentChange = dollarsChange / openPrice
            
            'instruct value to go to Summary Table
            Cells(summaryRow, 10).Value = Ticker
            
            Cells(summaryRow, 11).Value = dollarsChange
                'Change Color of the Cells: if dollarsChange is larger or equal to 0, change to green. if not, change to red
                If dollarsChange >= 0 Then
                    Cells(summaryRow, 11).Interior.ColorIndex = 4
                Else
                    Cells(summaryRow, 11).Interior.ColorIndex = 3
                End If
                
                    
            Cells(summaryRow, 12).Value = percentChange
                
            
            Cells(summaryRow, 13).Value = totalVolume
            
            'Increment summaryRow - instruct new summary value into the next Row
            summaryRow = summaryRow + 1
            
            
            'Reset the calculation on the next Row with next Ticker value
            openPrice = Cells(currentRow + 1, 3).Value
            
            'Reset totalVolume to 0
            totalVolume = 0
            
        
        End If
    
    
    Next currentRow
    
    'Go to next worksheet
    Next ws
    
    
End Sub

