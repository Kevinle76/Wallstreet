Sub Stockmarket()
    For Each ws In Worksheets
    
    
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearlychange"
        ws.Range("K1").Value = "Percentchange"
        ws.Range("L1").Value = "Totalstockvolume"
        ws.Range("N2").Value = "Greatest % Increase"
        ws.Range("N3").Value = "Greatest % Decrease"
        ws.Range("N4").Value = "Greatest Total Volume"
        ws.Range("O1").Value = "Ticker"
        ws.Range("P1").Value = "Value"

        Dim Tickername As String
        Dim LastRow As Long
        Dim Totalstockvolume As Double
        Dim Summarytablerow As Integer
        Dim Yearlyopen As Double
        Dim Yearlyclose As Double
        Dim Yearlychange As Double
        Dim Percentchange As Double
        Dim Greatestincrease As Double
        Dim Greatestdecrease As Double
        Dim Lastrowvalue As Double
        Dim Greatesttotalvolume As Double
         
        Totalstockvolume = 0
        Summarytablerow = 2
        Greatestincrease = 0
        Greatestdecrease = 0
        Greatesttotalvolume = 0
        
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        Yearlyopen = ws.Cells(2, 3)
        
        For i = 2 To LastRow
        
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1) Then
                
                    Tickername = ws.Cells(i, 1)
                    Totalstockvolume = Totalstockvolume + ws.Cells(i, 7)
                    
                    Yearlyclose = ws.Cells(i, 6)
                    Yearlychange = Yearlyclose - Yearlyopen
                        
                        If Yearlyopen = 0 Then
                            Percentchange = 0
                        
                        Else
                            Percentchange = Yearlychange / Yearlyopen * 100
                        
                        End If
            
            ws.Range("I" & Summarytablerow).Value = Tickername
            ws.Range("J" & Summarytablerow).Value = Yearlychange
            ws.Range("K" & Summarytablerow).Value = (Percentchange & "%")
            ws.Range("L" & Summarytablerow).Value = Totalstockvolume
            
            Totalstockvolume = 0
            Yearlyopen = ws.Cells(i + 1, 3)
    
                        If ws.Range("J" & Summarytablerow).Value >= 0 Then
                        
                             ws.Range("J" & Summarytablerow).Interior.ColorIndex = 4
            
                        Else
                        
                            ws.Range("J" & Summarytablerow).Interior.ColorIndex = 3
                        
                        End If
                 
            Summarytablerow = Summarytablerow + 1
            
            Else
                Totalstockvolume = Totalstockvolume + ws.Cells(i, 7)
         
         End If
         
         Next i
         
         LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
         
         
    For i = 2 To LastRow
         
        If ws.Range("K" & i).Value > ws.Range("P2").Value Then
        
               ws.Range("P2").Value = ws.Range("K" & i).Value
               ws.Range("O2").Value = ws.Range("I" & i).Value
               
        End If
            
        If ws.Range("K" & i).Value < ws.Range("P3").Value Then
                ws.Range("P3").Value = ws.Range("K" & i).Value
                ws.Range("O3").Value = ws.Range("I" & i).Value
        End If
            
        If ws.Range("L" & i).Value > ws.Range("P4").Value Then
                ws.Range("P4").Value = ws.Range("L" & i).Value
                ws.Range("O4").Value = ws.Range("I" & i).Value
        End If
                          
        Next i
             ws.Range("P2").NumberFormat = "0.00%"
             ws.Range("P3").NumberFormat = "0.00%"
                       
        Next ws
                
End Sub
                     
            
               
         




