Attribute VB_Name = "Module1"
Sub multipleYearStockData():

    ' to loop through every worksheet in the excel file
    For Each ws In Worksheets
    
        ' define all of my cells
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        ' define all my variables
        Dim wsName As String
        Dim i As Long
        Dim j As Long
        Dim tickerCount As Long
        Dim LastRowA As Long
        Dim LastRowI As Long
        Dim percentChange As Double
        Dim GreatestInc As Double
        Dim GreatestDec As Double
        Dim GreatestVol As Double
        
        wsName = ws.Name
        
        ' tell VBA where the last row of data is in the column and loop through it
        LastRowA = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        tickerCount = 2
        
        j = 2
            
            ' search through column from the bottom of the data up
            For i = 2 To LastRowA
            
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                ws.Cells(tickerCount, 9).Value = ws.Cells(i, 1).Value
                
                ws.Cells(tickerCount, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
                
                    If ws.Cells(tickerCount, 10).Value < 0 Then
                
                    ws.Cells(tickerCount, 10).Interior.ColorIndex = 3
                
                    Else
                
                    ws.Cells(tickerCount, 10).Interior.ColorIndex = 4
                
                    End If
                    
                    If ws.Cells(j, 3).Value <> 0 Then
                    percentChange = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
                    
                    ws.Cells(tickerCount, 11).Value = Format(percentChange, "Percent")
                    
                    Else
                    
                    ws.Cells(tickerCount, 11).Value = Format(0, "Percent")
                    
                    End If
                    
                ws.Cells(tickerCount, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
                
                tickerCount = tickerCount + 1
                
                j = i + 1
                
                End If
            
            Next i
            
        ' tell VBA where the last row of data is in the column and loop through it
        LastRowI = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        'Cell ranges for min's and max's
        GreatestVol = ws.Range("Q2").Value
        GreatestInc = ws.Range("K2").Value
        GreatestDec = ws.Range("K2").Value
        
            ' add in for loop
            For i = 2 To LastRowI
            
                ' these are conditions for max's and min's
                If ws.Cells(i, 12).Value > GreatestVol Then
                    GreatestVol = ws.Cells(i, 12).Value
                    ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                
                End If
                
                If ws.Cells(i, 11).Value > GreatestInc Then
                    GreatestInc = ws.Cells(i, 11).Value
                    ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                    
                End If
                
                If ws.Cells(i, 11).Value < GreatestDec Then
                    GreatestDec = ws.Cells(i, 11).Value
                    ws.Cells(3, 16).Value = ws.Cells(i, 9).Value

                End If
                
                ' displays the value for the coresponding variables into the correct format too
                ws.Range("Q2").Value = Format(GreatestInc, "Percent")
                ws.Range("Q3").Value = Format(GreatestDec, "Percent")
                ws.Range("Q4").Value = Format(GreatestVol, "Scientific")

                
            Next i
            
        'autofit all the columns so each cell has the correct spacing for it s value
        Worksheets(ws.Name).Columns("A:Z").AutoFit
            
    Next ws
    
End Sub


