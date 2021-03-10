Sub Ticker():

'Set Ws as Object Variable'
Dim Ws As Worksheet

'Loop through all of worksheets'
For Each Ws In Worksheets

'Header'
Ws.Range("I1").Value = "Ticker"
Ws.Range("J1").Value = "Yearly Change"
Ws.Range("K1").Value = "Percent Change"
Ws.Range("L1").Value = "Total Stock Volume"

'Declare TickerName & TickerColumn as Variable'
Dim TickerN As String
Dim TickerC As Long

'Ticker'
TickerN = " "
TickerC = 2

'Declare Opening and Closing Price as Variable'
Dim OpenP As Double
Dim ClosingP As Double

OpenP = 0
ClosingP = 0

'Column Count'
Dim Lastrow As Long
Dim i As Long

Lastrow = Ws.Cells(Rows.Count, 1).End(xlUp).Row

'Location for Opening Price'
OpenP = Ws.Cells(2, 3).Value

'Loop through Ticker Column to Last integer'
    For i = 2 To Lastrow
    
        'If Change in CellValue/Ticker in Worksheet then'
        If Ws.Cells(i + 1, 1).Value <> Ws.Cells(i, 1).Value Then
        
            'Input/Output to Column I and Declaring Column I for TickerN'
            TickerN = Ws.Cells(i, 1).Value
            
             ' Forumla for DeltaPrice and DeltaPercent'
                CloseP = Ws.Cells(i, 6).Value
                DeltaP = CloseP - OpenP
                
                ' Check Division by 0 condition
                If OpenP <> 0 Then
                    DeltaPercent = (DeltaP / OpenP) * 100
                End If
                
            'Total Ticker Volume'
            TtlTV = TtlTV + Ws.Cells(i, 7).Value
            
            'Print TickerN'
            Ws.Range("I" & TickerC).Value = TickerN
            
            'Print Delta P'
            Ws.Range("J" & TickerC).Value = DeltaP
                
                'DeltaPrice Green and Red colors'
                If (DeltaP > 0) Then
                    Ws.Range("J" & TickerC).Interior.ColorIndex = 4
                    
                ElseIf (DeltaP <= 0) Then
                    Ws.Range("J" & TickerC).Interior.ColorIndex = 3
                    
                End If
                
                ' Print the Ticker Name in the Summary Table, Column I
                Ws.Range("K" & TickerC).Value = (CStr(DeltaPercent) & "%")
                
                ' Print the Ticker Name in the Summary Table, Column J
                Ws.Range("L" & TickerC).Value = TtlTV
             
            'Add 1 to TickerC count'
            TickerC = TickerC + 1
            
            'Reset Delta'
            DeltaP = 0
            
        End If

    Next i
    
Next Ws

End Sub
