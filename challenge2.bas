Attribute VB_Name = "Módulo1"
Sub stock():

    Dim n, counter, j, open_date, close_date, a, b, p As Integer
    Dim open_stock, close_stock, yeardif, percent, vol, greatest, lowest, num1, num2, low1, low2 As Double
    Dim ws_num As Integer
    Dim ws As Worksheet
    Dim ticker As String
    ws_num = ThisWorkbook.Worksheets.Count
    
    For i = 1 To ws_num
    
    ThisWorkbook.Worksheets(i).Activate
    
        n = Worksheets(i).UsedRange.Rows.Count
        
        j = 2
        
        ticker = Cells(2, 1).Value
        open_date = 2
        open_stock = Cells(open_date, 3).Value
        
        'headers
        
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly change"
        Cells(1, 11).Value = "Percent"
        Cells(1, 12).Value = "Total stock volume"
        
        Cells(2, 15).Value = "Greatest % percent change"
        Cells(3, 15).Value = "Greatest decrease change"
        Cells(4, 15).Value = "Greatest total volume"
        Cells(1, 16).Value = "ticker"
        Cells(1, 17).Value = "value"
        
        'calculator
        
        For counter = 2 To n
            
            
            If Cells(counter + 1, 1).Value <> ticker Then
                ticker = Cells(counter, 1).Value
                open_stock = Cells(open_date, 3).Value
                vol = vol + Cells(counter, 7).Value
                Cells(j, 9).Value = ticker
                Cells(j, 12).Value = vol
                ticker = Cells(counter + 1, 1).Value
                open_date = counter + 1
                close_date = counter
                close_stock = Cells(close_date, 6).Value
                yeardif = close_stock - open_stock
                percent = yeardif / open_stock
                open_stock = Cells(open_date, 3).Value
                Cells(j, 10).Value = yeardif
                Cells(j, 11).Value = percent
                Range("K" & j).NumberFormat = "0.00%"
                Range("J" & j).NumberFormat = "0.00"
                j = j + 1
                vol = 0
                
                endrow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "J").End(xlUp).Row
                If Range("J" & endrow).Value < 0 Then
                   Range("J" & endrow).Interior.ColorIndex = 3
                Else
                    Range("J" & endrow).Interior.ColorIndex = 4
                End If
            Else
                vol = vol + Cells(counter, 7).Value
                
            End If
            
            'conditional color
            
            
            
            
                    
                        
        Next counter
        
        
        
            Range("A:Q").Columns.AutoFit
            num1 = Cells(2, 11).Value
            ticker = Cells(2, 9).Value
        'greatest change
        
        For a = 3 To n
            num2 = Cells(a, 11).Value
                    
            If num2 > num1 Then
                num1 = num2
                ticker = Cells(a, 9).Value
            End If
                       
        Next a
            greatest = num1
            Cells(2, 17).Value = greatest
            Cells(2, 17).NumberFormat = "0.00%"
            Cells(2, 16).Value = ticker
        'lowest change
        
            low1 = Cells(2, 11).Value
            ticker = Cells(2, 9).Value
        For a = 3 To n
            low2 = Cells(a, 11).Value
                    
            If low2 < low1 Then
                low1 = low2
                ticker = Cells(a, 9).Value
            End If
                       
        Next a
            lowest = low1
            Cells(3, 17).Value = lowest
            Cells(3, 17).NumberFormat = "0.00%"
            Cells(3, 16).Value = ticker
             
            'greatest volume
            
            num1 = Cells(2, 12).Value
            ticker = Cells(2, 9).Value
        For a = 3 To n
            num2 = Cells(a, 12).Value
                    
            If num2 > num1 Then
                num1 = num2
                ticker = Cells(a, 9).Value
            End If
                       
        Next a
            greatest = num1
            Cells(4, 17).Value = greatest
            Cells(4, 16).Value = ticker
            Cells(4, 17).NumberFormat = "0"
    Next i
    
    
End Sub

