Attribute VB_Name = "Module1"
Sub StockTicker()

    'Variable counting the rows
    Dim lastrow As Double
    Dim SummaryLastRow As Double
    
    'Variable for holding Stock Ticker
    Dim Ticker As String
    
    Dim YearlyChng As Double
    Dim PercentChng As Double
    Dim TotalVolume As Double
    Dim SummaryTable As Integer
    
    Dim BeginingDate As Double
    Dim EndingDate As Double
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    
    'Dim WorksheetName As String
                    
    For Each ws In Worksheets
    
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percentage Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("P1").Value = ws.Range("I1").Value
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volumn"
                           
        'find the last row and assing that row number to the variable
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
          
        'defining some starting variables
        WorksheetName = ws.Name
        TotalVolume = 0
        OpenPrice = 0
        SummaryTable = 2
        BeginingDate = 30001231
        EndingDate = 20000101
    
        'Check Ticker name to make sure other values come from same ticker
        For Row = 2 To lastrow
            
            If ws.Cells(Row + 1, 1).Value <> ws.Cells(Row, 1).Value Then
            
                Ticker = ws.Cells(Row, 1).Value
                
                TotalVolume = TotalVolume + ws.Cells(Row, 7).Value
                
                ClosePrice = ws.Cells(Row, 6).Value
                
                YearlyChng = (ClosePrice - OpenPrice)
                
                'PercentChng = ((YearlyChng / OpenPrice) * 100)
                PercentChng = (YearlyChng / OpenPrice)
            
                ws.Range("I" & SummaryTable).Value = Ticker
                
                ws.Range("J" & SummaryTable).Value = YearlyChng
                
                If ws.Range("J" & SummaryTable).Value < 0 Then
                
                    ws.Range("J" & SummaryTable).Interior.ColorIndex = 3
                        
                ElseIf ws.Range("J" & SummaryTable) > 0 Then
                    
                    ws.Range("J" & SummaryTable).Interior.ColorIndex = 4
                    
                Else
                    
                    ws.Range("J" & SummaryTable).Interior.ColorIndex = 2
                    
                End If
                
                ws.Range("K" & SummaryTable).NumberFormat = "0.00%"
                
                ws.Range("K" & SummaryTable).Value = PercentChng
                
                ws.Range("L" & SummaryTable).NumberFormat = "#,##0"
                
                ws.Range("L" & SummaryTable).Value = TotalVolume
                
                SummaryTable = SummaryTable + 1
                      
                TotalVolume = 0
                OpenPrice = 0
            
            Else
                
                TotalVolume = TotalVolume + ws.Cells(Row, 7).Value
                
                If OpenPrice = 0 Then
                    OpenPrice = ws.Cells(Row, 3).Value
                End If
                                          
            End If
            
            'Getting the first date in the for loop
            'If BeginingDate > Cells(Row, 2).Value Then
                'BeginingDate = Cells(Row, 2).Value
           ' End If
            
            'Getting the last date in the for loop
            'If EndingDate < Cells(Row, 2).Value Then
                'EndingDate = Cells(Row, 2).Value
            'End If
            
            'YearlyChg = ClosePrice  '- OpenPrice)
            
            'Cells(Row, 10).Value = YearlyChng
            
        Next Row
              
        Dim SummaryTicker As String
        Dim GreatPerInc As Double
        Dim GreatPerDec As Double
        Dim GreatTotVal As Double
              
        GreatPerInc = ws.Cells(2, 11).Value
        GreatPerDec = ws.Cells(2, 11).Value
        GreatTotVal = ws.Cells(2, 12).Value
            'MsgBox (GreatPerInc & vbNewLine & GreatPerDec & vbNewLine & GreatTotVal)
            
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        ws.Range("Q4").NumberFormat = "#,##0"
        
        SummaryLastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
            'MsgBox (SummaryLastRow & " Is the last row of the summar column")
            
        For SumRow = 2 To SummaryLastRow
        
            If GreatPerInc < ws.Cells(SumRow, 11).Value Then
                GreatPerInc = ws.Cells(SumRow, 11).Value
                SummaryTicker = ws.Cells(SumRow, 9).Value
            End If
            
        Next SumRow
        
        ws.Range("Q2") = GreatPerInc
        ws.Range("P2") = SummaryTicker
        SummaryTicker = ""
            
        'SummaryLastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
            
        For SumRow = 2 To SummaryLastRow
        
            If GreatPerDec >= ws.Cells(SumRow, 11).Value Then
                GreatPerDec = ws.Cells(SumRow, 11).Value
                SummaryTicker = ws.Cells(SumRow, 9).Value
            End If
            
        Next SumRow
            
        ws.Range("Q3") = GreatPerDec
        ws.Range("P3") = SummaryTicker
        SummaryTicker = ""
    
        'SummaryLastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
            
        For SumRow = 2 To SummaryLastRow
        
            If GreatTotVal < ws.Cells(SumRow, 12).Value Then
                GreatTotVal = ws.Cells(SumRow, 12).Value
                SummaryTicker = ws.Cells(SumRow, 9).Value
            End If
            
        Next SumRow
            
        ws.Range("Q4") = GreatTotVal
        ws.Range("P4") = SummaryTicker
        ws.Columns("I:Q").EntireColumn.AutoFit
        SummaryTicker = ""
        
        'Range("L1.Q1").EntireColumn.AutoFit
        'Range("Q1").EntireColumn.AutoFit
                
            'MsgBox (BeginingDate & "  " & EndingDate)
            'MsgBox (ClosePrice)
            ''Ticker = Cells(Row, 1).Value
            ''For Col = 1 To 7
        
        'MsgBox (WorksheetName & " worksheet")
        
    Next ws
        
        
End Sub

