Attribute VB_Name = "Module1"
Sub stock_data():


For Each ws In Worksheets


    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    
    
    Dim openvalue As Double
    Dim closevalue As Double
    Dim yearlychange As Double
    Dim totalvolume As Double
    Dim rowcount As Integer
    
    
    openvalue = ws.Cells(2, 3).Value
    closevalue = 0
    yearlychange = 0
    totalvolume = 0
    rowcount = 2
    
    
    Dim lastrow As Long
    
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To lastrow
    
        If (openvalue = 0) Then
        openvalue = ws.Cells(i, 3).Value
        End If
    
        If (ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value) Then
        totalvolume = totalvolume + ws.Cells(i, 7).Value
              
        Else
        totalvolume = totalvolume + ws.Cells(i, 7).Value
        closevalue = ws.Cells(i, 6).Value
        yearlychange = closevalue - openvalue
                
        ws.Range("I" & rowcount).Value = ws.Cells(i, 1).Value
        ws.Range("J" & rowcount).Value = yearlychange
        ws.Range("L" & rowcount).Value = totalvolume
                
            If (openvalue <> 0) Then
            ws.Range("K" & rowcount).Value = yearlychange / openvalue
            Else: ws.Range("K" & rowcount).Value = 0
            End If
                
        openvalue = ws.Cells(i + 1, 3).Value
        closevalue = 0
        yearlychange = 0
        totalvolume = 0
        rowcount = rowcount + 1
        
        End If
    
    Next i
    
    
    Dim greatest_up As Double
    Dim ticker1 As String
    Dim greatest_down As Double
    Dim ticker2 As String
    Dim greatest_volume As Double
    Dim ticker3 As String
    
    greatest_up = ws.Range("K2").Value
    greatest_down = ws.Range("K2").Value
    greatest_volume = ws.Range("L2").Value
    
    
    
    For i = 2 To rowcount - 1
    
        If (ws.Range("K" & i).Value > greatest_up) Then
        greatest_up = ws.Range("K" & i).Value
        ticker1 = ws.Range("I" & i).Value
        
        ElseIf (ws.Range("K" & i).Value < greatest_down) Then
        greatest_down = ws.Range("K" & i).Value
        ticker2 = ws.Range("I" & i).Value
        
        End If
        
    Next i
    
    
    
    For i = 2 To rowcount - 1
        If ws.Range("L" & i).Value > greatest_volume Then
        greatest_volume = ws.Range("L" & i).Value
        ticker3 = ws.Range("I" & i).Value
        End If
    Next i
     
     
     
    ws.Range("P2").Value = ticker1
    ws.Range("P3").Value = ticker2
    ws.Range("P4").Value = ticker3
    ws.Range("Q2").Value = greatest_up
    ws.Range("Q3").Value = greatest_down
    ws.Range("Q4").Value = greatest_volume
    
    
    ws.Columns("I:Q").AutoFit
    ws.Range("K2:K" & rowcount - 1).NumberFormat = "0.00%"
    ws.Range("Q2:Q3").NumberFormat = "0.00%"
    
    
    
    For i = 2 To rowcount - 1
        If (ws.Range("J" & i).Value > 0) Then
        ws.Range("J" & i).Interior.ColorIndex = 4
        Else: ws.Range("J" & i).Interior.ColorIndex = 3
        End If
    Next i



Next ws


End Sub
