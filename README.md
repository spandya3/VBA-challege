# VBA-challege

'VBA Code
Sub stock_price()

    Dim qchange As Double
    Dim changep As Double
    Dim totalVolume As Double
    Dim W As Worksheet
    Dim start As Long
    Dim columnIndex As Long
    Dim rowIndex As Long
    Dim rowCount As Long
    Dim increase As Double
    Dim decrease As Double
    Dim volnum As Long

  

    For Each W In Worksheets
        
        ' Clear existing data in output columns
        W.Range("I2:L" & W.Cells(W.Rows.Count, "I").End(xlUp).Row).ClearContents
        W.Range("P3:P5").ClearContents
        W.Range("Q3:Q5").ClearContents
        
        ' Initialize variables
        totalVolume = 0
        changep = 0
        start = 2
        columnIndex = 0
        rowCount = W.Cells(W.Rows.Count, "A").End(xlUp).Row
        
        W.Range("I1").Value = "Ticker"
        W.Range("J1").Value = "Quarterly Change"
        W.Range("K1").Value = "Percent Change"
        W.Range("L1").Value = "Total Stock Volume"
        W.Range("O3").Value = "Greatest Percent % Increase"
        W.Range("O4").Value = "Greatest Percent % Decrease"
        W.Range("O5").Value = "Greatest Total Volume"
        W.Range("P2").Value = "Ticker"
    W.Range("Q2").Value = "Value"
        
        ' Loop through rows
        For rowIndex = 2 To rowCount
        
            If W.Cells(rowIndex + 1, 1).Value <> W.Cells(rowIndex, 1).Value Then
                
                totalVolume = totalVolume + W.Cells(rowIndex, 7).Value
                
                If totalVolume = 0 Then
                    W.Cells(2 + columnIndex, "I").Value = W.Cells(rowIndex, 1).Value
                    W.Cells(2 + columnIndex, "J").Value = 0
                    W.Cells(2 + columnIndex, "K").Value = "%" & 0
                    W.Cells(2 + columnIndex, "L").Value = 0
                Else
                    If W.Cells(start, 3).Value = 0 Then
                        For Search = start To rowIndex
                            If W.Cells(Search, 3).Value <> 0 Then
                                start = Search
                                Exit For
                            End If
                        Next Search
                    End If
                    
                    qchange = W.Cells(rowIndex, 6).Value - W.Cells(start, 3).Value
                    If W.Cells(start, 3).Value <> 0 Then
                        changep = qchange / W.Cells(start, 3).Value
                    Else
                        changep = 0
                    End If
                    
                    W.Cells(2 + columnIndex, "I").Value = W.Cells(rowIndex, 1).Value
                    W.Cells(2 + columnIndex, "J").Value = qchange
                    W.Cells(2 + columnIndex, "J").NumberFormat = "0.00"
                    W.Cells(2 + columnIndex, "K").Value = "%" & changep * 100
                    W.Cells(2 + columnIndex, "K").NumberFormat = "0.00%"
                    W.Cells(2 + columnIndex, "L").Value = totalVolume
                    
                    ' Color formatting based on change percent
                    Select Case changep
                        Case Is > 0
                            W.Cells(2 + columnIndex, "J").Interior.ColorIndex = 4
                        Case Is < 0
                            W.Cells(2 + columnIndex, "J").Interior.ColorIndex = 3
                        Case Else
                            W.Cells(2 + columnIndex, "J").Interior.ColorIndex = 0
                    End Select
                End If
                
                ' Reset variables for next iteration
                totalVolume = 0
                qchange = 0
                columnIndex = columnIndex + 1
                start = rowIndex + 1
                
            Else
                totalVolume = totalVolume + W.Cells(rowIndex, 7).Value
            End If
            
        Next rowIndex
        
        ' Find and display greatest percent increase, decrease, and total volume
        If rowCount > 1 Then
            increase = Application.WorksheetFunction.Match(Application.WorksheetFunction.Max(W.Range("K2:K" & rowCount)), W.Range("K2:K" & rowCount), 0)
            decrease = Application.WorksheetFunction.Match(Application.WorksheetFunction.Min(W.Range("K2:K" & rowCount)), W.Range("K2:K" & rowCount), 0)
            volnum = Application.WorksheetFunction.Match(Application.WorksheetFunction.Max(W.Range("L2:L" & rowCount)), W.Range("L2:L" & rowCount), 0)
            
            W.Range("P3").Value = W.Cells(increase + 1, "I").Value
            W.Range("P4").Value = W.Cells(decrease + 1, "I").Value
            W.Range("P5").Value = W.Cells(volnum + 1, "I").Value
            
            W.Range("Q3") = "%" & WorksheetFunction.Max(W.Range("K2: K" & rowCount)) * 100
            W.Range("Q4") = "%" & WorksheetFunction.Min(W.Range("K2: K" & rowCount)) * 100
            W.Range("Q5") = WorksheetFunction.Max(W.Range("L2: L" & rowCount))
         
          
        End If
        
    Next W


End Sub

