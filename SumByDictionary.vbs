Sub MainSum()
    'This runs all the separate functions
    Dim dict As Dictionary
    
    Set dict = AllSums
    
    WriteToImmediate dict
    
    WriteToWorksheet dict, ThisWorkbook.ActiveSheet

End Sub

Private Function AllSums() As Dictionary
    'This creates the dictionary using classes defined in a separate module
    Dim dict As New Dictionary
    Dim sh As Worksheet
    Set sh = ThisWorkbook.ActiveSheet
    Dim rng As Range
    Set rng = sh.Range("A1").CurrentRegion

    Dim oStock As clsStock, i As Long, ticker As String

For i = 2 To rng.Rows.Count
        ticker = rng.Cells(i, 1).Value
    
    If dict.Exists(ticker) = True Then
            Set oStock = dict(ticker)
        Else
            Set oStock = New clsStock
            dict.Add ticker, oStock
        End If
    
    oStock.open_price = oStock.open_price + rng.Cells(i, 3).Value
    oStock.high_price = oStock.high_price + rng.Cells(i, 4).Value
    oStock.low_price = oStock.low_price + rng.Cells(i, 5).Value
    oStock.close_price = oStock.close_price + rng.Cells(i, 6).Value
    oStock.volume = oStock.volume + rng.Cells(i, 7).Value
   Next i
   Set AllSums = dict
End Function

Private Sub WriteToImmediate(dict As Dictionary)
    'Uses debug window to check results
    Dim key As Variant, oStock As clsStock
    
    For Each key In dict.Keys
        Set oStock = dict(key)
        With oStock
            Debug.Print key, .open_price, .high_price, .low_price, .close_price, .volume
        End With
        
    Next key
    
End Sub
Private Sub WriteToWorksheet(dict As Dictionary, sh As Worksheet)
    
    'Finally this writes the results to the active sheet and formats the change
    'Also calculates the yearly change and percentage change
    Dim row As Long
    row = 2
    
    Dim key As Variant, oStock As clsStock
            sh.Cells(1, 9).Value = "Ticker"
            sh.Cells(1, 10).Value = "Total Opening Price"
            sh.Cells(1, 11).Value = "Total High Price"
            sh.Cells(1, 12).Value = "Total Low Price"
            sh.Cells(1, 13).Value = "Total Closing Price"
            sh.Cells(1, 14).Value = "Total Volume"
            sh.Cells(1, 15).Value = "Total Price Change"
            sh.Cells(1, 16).Value = "Percent Change"
            
            ' Read through the dictionary
    
            
    For Each key In dict.Keys
        Set oStock = dict(key)
        With oStock
            ' Write out the values
            
            sh.Cells(row, 9).Value = key
            sh.Cells(row, 10).Value = .open_price
            sh.Cells(row, 11).Value = .high_price
            sh.Cells(row, 12).Value = .low_price
            sh.Cells(row, 13).Value = .close_price
            sh.Cells(row, 14).Value = .volume
            sh.Cells(row, 15).Value = .close_price - .open_price
            sh.Cells(row, 16).Value = ((.close_price - .open_price) / .open_price)
                If sh.Cells(row, 15).Value < 0 Then
                    sh.Cells(row, 15).Interior.ColorIndex = 3
                    Else: sh.Cells(row, 15).Interior.ColorIndex = 4
                End If
            row = row + 1
        End With
        
    Next key
    
End Sub
