Attribute VB_Name = "Module1"
Sub Challenge()
Dim i As Long
Dim volume As Double
volume = 0
Dim ws As Worksheet
Dim LastRow As Long
Dim Ticker As String
Dim OpenPrice As Double
Dim ClosePrice As Double
Dim ChangePrice As Double
Dim ChangePricepct As Double
Dim OutputRow As Long
Dim maxpct As Double
Dim minpct As Double
Dim maxtotal As Double


For Each ws In ThisWorkbook.Worksheets
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    OutputRow = 2
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greastest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
   

    
    For i = 2 To LastRow

        If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
            'Start new group'
            
            Ticker = ws.Cells(i, 1).Value
            volume = ws.Cells(i, 7).Value
            OpenPrice = ws.Cells(i, 3).Value
            
            
        ElseIf ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            'End of group'
            
            volume = volume + ws.Cells(i, 7).Value
            ClosePrice = ws.Cells(i, 6).Value
            ChangePrice = ClosePrice - OpenPrice
            ChangePricepct = ChangePrice / OpenPrice
            
            ws.Cells(OutputRow, 9).Value = Ticker
            
            ws.Cells(OutputRow, 10).Value = ChangePrice
            If ws.Cells(OutputRow, 10).Value > 0 Then
                ws.Cells(OutputRow, 10).Interior.Color = vbGreen
                ElseIf ws.Cells(OutputRow, 10).Value < 0 Then
                ws.Cells(OutputRow, 10).Interior.Color = vbRed
            End If
              
                ws.Cells(OutputRow, 11).Value = ChangePricepct
                ws.Columns("K:K").NumberFormat = "0.00%"
                ws.Cells(OutputRow, 12).Value = volume
                OutputRow = OutputRow + 1
                volume = 0
            
            
        Else
            'Middle of group'
            volume = volume + ws.Cells(i, 7).Value
            
            
        End If
        
        
    Next i
    
    maxpct = ws.Cells(2, 11).Value
    minpct = ws.Cells(2, 11).Value
    maxtotal = ws.Cells(2, 12).Value

    
    For j = 2 To LastRow
        If ws.Cells(j, 11) > maxpct Then
            maxpct = ws.Cells(j, 11).Value
            ws.Cells(2, 16).Value = ws.Cells(j, 9).Value
            ws.Cells(2, 17).Value = maxpct
            ws.Cells(2, 17).NumberFormat = "0.00%"
        End If
    
        
        If ws.Cells(j, 11) <= minpct Then
            minpct = ws.Cells(j, 11).Value
            ws.Cells(3, 16).Value = ws.Cells(j, 9).Value
            ws.Cells(3, 17).Value = minpct
            ws.Cells(3, 17).NumberFormat = "0.00%"
        End If
        
        If ws.Cells(j, 12).Value >= maxtotal Then
            maxtotal = ws.Cells(j, 12).Value
            ws.Cells(4, 17).Value = maxtotal
            ws.Cells(4, 16).Value = ws.Cells(j, 9).Value
        End If
        
    Next j
     ws.Columns("I:Q").AutoFit
    

Next ws


End Sub

