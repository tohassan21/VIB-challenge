Attribute VB_Name = "Module1"
Option Explicit

Sub quarterloop()

   'counter variable
   Dim row As Long
   'stock name variable
   Dim ticker As String
   'opening price in beginning of quarter variable
   Dim quarteropen As Double
   'closing price in end of quarter variable
   Dim quarterclose As Double
   'difference between quarterclose and quarteropen variable
   Dim quarterlychange As Double
   'percent change
   Dim percentchange As Double
   
   'total volume for each unique stock variable
   Dim totalstockvolume As Double
   
   quarterlychange = 0
   totalstockvolume = 0
   quarteropen = 0
   quarterclose = 0
   
   'keep track of each unique stock in a summary table
   Dim summaryrow As Long
   summaryrow = 2
   
   Dim ws As Worksheet
   
   'loop through each worksheet
   For Each ws In Worksheets
   
        'last row variable
        Dim lastrow As Long
        'determine last row
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).row
        
        'QUARTER OPEN STUFF
        'set the quarteropen price
        quarteropen = ws.Cells(2, 3).Value

       'loop through all stocks
       For row = 2 To lastrow
                
            'check if we are still within the same stock, if not
            If ws.Cells(row + 1, 1).Value <> ws.Cells(row, 1).Value Then
                'set the ticker name
                ticker = ws.Cells(row, 1).Value
                'add to the total stock volume
                totalstockvolume = totalstockvolume + ws.Cells(row, 7).Value
                
                'QUARTER CLOSE STUFF
                'set the quarterclose price
                quarterclose = ws.Cells(row, 6).Value
                
                'QUARTERLY CHANGE STUFF
                'calculate quarterly change
                quarterlychange = quarterclose - quarteropen
                'print quarter change
                ws.Range("J" & summaryrow).Value = quarterlychange
                'calculate percent change
                percentchange = ((quarterclose - quarteropen) / quarteropen)
                'print percent change
                ws.Range("K" & summaryrow).Value = percentchange
                'format percent change as percentage
                ws.Range("K" & summaryrow).NumberFormat = "0.00%"
                
                
                'Color change based on quarterly change positive or negative
                If ws.Range("J" & summaryrow).Value > 0 Then
                    ' Set the Cell Colors to Red
                    ws.Range("J" & summaryrow).Interior.ColorIndex = 4
                ElseIf ws.Range("J" & summaryrow).Value < 0 Then
                    ' Set the Cell Colors to Green
                    ws.Range("J" & summaryrow).Interior.ColorIndex = 3
                End If
                
                'print ticker name in summary table
                ws.Range("I" & summaryrow).Value = ticker
                
                'print total stock volume to summary table
                ws.Range("L" & summaryrow).Value = totalstockvolume
                'add one to summary table row
                summaryrow = summaryrow + 1
                'reset total stock volume
                totalstockvolume = 0
                
                'QUARTER OPEN STUFF
                'set the quarteropen price
                quarteropen = ws.Cells(row + 1, 3).Value
                
            'if the cell immediately following a row is the same as the previous row
            Else
                'add to the total stock volume
                totalstockvolume = totalstockvolume + ws.Cells(row, 7).Value
            End If
        
        Next row
        
        
        'summary last row variable
        Dim summarylastrow As Long
        'determine summary last row
        summarylastrow = ws.Cells(Rows.Count, 9).End(xlUp).row
        
        Dim summaryticker As String
        Dim maxincrease As Double
        Dim maxdecrease As Double
        Dim maxtotalvolume As Double
        
        For summaryrow = 2 To summarylastrow
            'find max percent increase
            If ws.Cells(summaryrow, 11).Value > maxincrease Then
                maxincrease = ws.Cells(summaryrow, 11).Value
                summaryticker = ws.Cells(summaryrow, 9).Value
                ' Populate the summary table with the ticker and value
                ws.Range("P2").Value = summaryticker
                ws.Range("Q2").Value = maxincrease
                ws.Range("Q2").NumberFormat = "0.00%"
            End If
            
            'find max percent decrease
             If ws.Cells(summaryrow, 11).Value < maxdecrease Then
                maxdecrease = ws.Cells(summaryrow, 11).Value
                summaryticker = ws.Cells(summaryrow, 9).Value
                ws.Range("P3").Value = summaryticker
                ws.Range("Q3").Value = maxdecrease
                ws.Range("Q3").NumberFormat = "0.00%"
            End If
            
            'find max total volume
             If ws.Cells(summaryrow, 12).Value > maxtotalvolume Then
                maxtotalvolume = ws.Cells(summaryrow, 12).Value
                summaryticker = ws.Cells(summaryrow, 9).Value
                ws.Range("P4").Value = summaryticker
                ws.Range("Q4").Value = maxtotalvolume
            End If
        Next summaryrow
                

        'LABELS
        'summary table labels
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Quarterly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        'additional summary table labels
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        'RESET VALUES BEFORE NEXT SHEET
        summaryrow = 2
        quarterlychange = 0
        totalstockvolume = 0
        quarteropen = 0
        quarterclose = 0
        row = 2

        maxincrease = 0
        maxdecrease = 0
        maxtotalvolume = 0
        
    Next ws
        
End Sub
