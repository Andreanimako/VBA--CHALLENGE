Attribute VB_Name = "Module1"
Sub Multistockmarket()

    For Each ws In Worksheets
    
        Dim WSname As String
        Dim tickername As String
        Dim yearlychange As Double
        Dim perentchange As Double
        Dim totalstock As Double
        Dim summarytablerow As Integer
        Dim lastrow As Long
        Dim openprice As Double
        Dim closeprice As Double
        Dim counter As Integer
        
        
        
        
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        WSname = ws.Name
        summarytablerow = 2
        totalstock = 0
        counter = 0
        
        
        
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % increase"
        ws.Cells(3, 15).Value = "Greatest % decrease"
        ws.Cells(4, 15).Value = "Greatest total volume"
        
        For i = 2 To lastrow
        
            counter = counter + 1
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
               
                
                ticker = ws.Cells(i, 1).Value
                
                openprice = ws.Cells(((i + 1) - counter), 3).Value
                closeprice = ws.Cells(i, 6).Value
                
                yearlychange = closeprice - openprice
                
                percentchange = (yearlychange / openprice) * 100 & "%"
        
                
                totalstock = totalstock + ws.Cells(i, 7).Value
                
                
                ws.Range("I" & summarytablerow).Value = ticker
                
                ws.Range("L" & summarytablerow).Value = totalstock
                
                ws.Range("K" & summarytablerow).Value = percentchange
                
                ws.Range("j" & summarytablerow).Value = yearlychange
                
                summarytablerow = summarytablerow + 1
                
                totalstock = 0
                yearlychange = 0
                counter = 0
                
            
            Else
                 totalstock = totalstock + ws.Cells(i, 7).Value
                
            
            End If
            
        Next i
        
        For i = 2 To 3001
        
            If ws.Cells(i, 11).Value = WorksheetFunction.Max(ws.Range("K1:K3001")) Then
                ws.Cells(2, 17).Value = (ws.Cells(i, 11).Value) * 100 & "%"
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                
            ElseIf ws.Cells(i, 11).Value = WorksheetFunction.Min(ws.Range("K1:K3001")) Then
                ws.Cells(3, 17).Value = (ws.Cells(i, 11).Value) * 100 & "%"
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
        
            ElseIf ws.Cells(i, 12).Value = WorksheetFunction.Max(ws.Range("L1:L3001")) Then
                ws.Cells(4, 17).Value = ws.Cells(i, 12).Value
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                
                
            End If
           
        Next i
        
        For i = 2 To 3001
        
            If ws.Cells(i, 10).Value > 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 4
            Else
                ws.Cells(i, 10).Interior.ColorIndex = 3
            
            End If
            
        Next i
        
    Next ws
    
    End Sub

