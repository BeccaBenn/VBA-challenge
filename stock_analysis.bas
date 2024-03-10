Attribute VB_Name = "Module1"
'Attribute VB_Name = "Module1"
Sub Stock_Analysis()

    ' Loop through worksheets
    For Each ws In ThisWorkbook.Worksheets
        ' Assigns variables
        Dim ticker As String
        Dim year_change As Double
        Dim percent_change As Double
        Dim total_volume As Double
        Dim new_volume As Double
        Dim myRange As Long
        Dim count As Long
        Dim opening As Double
        Dim closing As Double
        Dim sum_range As Long
        
        
        ' Assign column headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        ' Finds last row of data in worksheet
        myRange = ws.Cells(Rows.count, 1).End(xlUp).Row
        
        ' Counter starts at row 2
        count = 2
        
        ' Loops through all rows
        For i = 2 To myRange
        
            ' Stores ticker
            ticker = ws.Cells(i, 1).Value
            
            If ws.Cells(i - 1, 1).Value <> ticker Then
            
                ' Stores opening price
                opening = ws.Cells(i, 3).Value
            End If
            
            If ws.Cells(i, 1).Value = ticker Then
            
                ' Stores value of each row's volume and adds to total volume
                new_volume = ws.Cells(i, 7).Value
                total_volume = total_volume + new_volume
                
                ' Sends total volume to final summary columns
                ws.Cells(count, 12).Value = total_volume
            End If
            
            ' Checks if ticker changes
            If ws.Cells(i + 1, 1).Value <> ticker Then
                ws.Cells(count, 10).Value = ticker
                
                ' Resets total volume
                total_volume = 0
                
                ' Stores closing price
                closing = ws.Cells(i, 6).Value

                ' Finds yearly change
                year_change = closing - opening
                
            
                ' Finds percent change
                If opening <> closing Then
                    percent_change = (year_change / opening) * 1
                Else
                    percent_change = 0
                End If
                
                               
                ' Adds ticker, yearly change, and percent change to final summary columns
                ws.Cells(count, 9).Value = ticker
                ws.Cells(count, 10).Value = year_change
                ws.Cells(count, 11).Value = percent_change
                
                ' Formats percent change
                ws.Cells(count, 11).NumberFormat = "0.00%"
                
                ' Conditional formatting for yearly changes
                If year_change > 0 Then
                    ws.Cells(count, 10).Interior.Color = RGB(0, 255, 0)
                ElseIf year_change < 0 Then
                    ws.Cells(count, 10).Interior.Color = RGB(255, 0, 0)
                End If
                
                ' Increments the summary rows
                count = count + 1
                
            End If
            
        Next i
        
        ' Finds last row of data in the summary columns
        sum_range = ws.Cells(Rows.count, 9).End(xlUp).Row
        For i = 2 To sum_range
            'Finds greatest percent increase
            If ws.Cells(i, 11).Value > ws.Range("Q2").Value Then
                ws.Range("Q2").Value = ws.Cells(i, 11).Value
                ws.Range("P2").Value = ws.Cells(i, 9).Value
                ws.Range("Q2").NumberFormat = "0.00%"
                
            'Finds greatest percent decrease
            ElseIf ws.Cells(i, 11).Value < ws.Range("Q3").Value Then
                ws.Range("Q3").Value = ws.Cells(i, 11).Value
                ws.Range("P3").Value = ws.Cells(i, 9).Value
                ws.Range("Q3").NumberFormat = "0.00%"
            End If
            
            'Finds greatest total volume
            If ws.Cells(i, 12).Value > ws.Range("Q4").Value Then
                ws.Range("Q4").Value = ws.Cells(i, 12).Value
                ws.Range("P4").Value = ws.Cells(i, 9).Value
            End If
            
        Next i
        
        'Adjusts column width
        ws.Range("I:Q").Columns.AutoFit
    Next
    
End Sub


