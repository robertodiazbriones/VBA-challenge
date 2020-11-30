Attribute VB_Name = "Module1"
Sub hwtesting()

Dim a As Integer
Dim b As Long
Dim c As Integer
Dim d As Integer
Dim e As Integer
Dim lastws As Integer
Dim nws As Integer
Dim Ticker As String
Dim TotalSV As LongLong
Dim lastrow As Long
Dim lastrow2 As Long
Dim CloseValue As Double
Dim YearlyChange As Double
Dim PercentChar As Double
Dim GreatInc As Double
Dim GreatDec As Double
Dim GreatAux As Double
Dim StockVol As LongLong
Dim StockVolAux As LongLong
Dim GreatIncTicker As String
Dim GreatDecTicker As String
Dim StockVolTicker As String



a = 1
nws = Application.Sheets.Count

For a = 1 To nws

    Worksheets(a).Range("I1").Value = "Ticker"
    Worksheets(a).Range("J1").Value = "Yearly Change"
    Worksheets(a).Range("K1").Value = "Percent Char"
    Worksheets(a).Range("L1").Value = "Total Stock Volume"
    Worksheets(a).Range("N2").Value = "Greatest % Increase"
    Worksheets(a).Range("N3").Value = "Greatest % Decrease"
    Worksheets(a).Range("N4").Value = "Greatest Total Volume"
    
    lastrow = Worksheets(a).Cells(Rows.Count, 1).End(xlUp).Row + 1
    
    
    
    
    c = 1
    TotalSV = 0
    GreatInc = 0
    GreatDec = 0
    StockVol = 0
    Ticker = Worksheets(a).Cells(2, 1).Value
    OpenValue = Worksheets(a).Cells(2, 3).Value
   
        For b = 2 To lastrow
            If Ticker <> Worksheets(a).Cells(b, 1).Value Then
                c = c + 1
                'Print Values for last ticker
                
                Worksheets(a).Cells(c, 9).Value = Ticker
                Worksheets(a).Cells(c, 12).Value = TotalSV
                YearlyChange = CloseValue - OpenValue
                Worksheets(a).Cells(c, 10).Value = YearlyChange
                If Worksheets(a).Cells(c, 10).Value > 0 Then
                    Worksheets(a).Cells(c, 10).Interior.ColorIndex = 4
                ElseIf Worksheets(a).Cells(c, 10).Value < 0 Then
                    Worksheets(a).Cells(c, 10).Interior.ColorIndex = 3
                End If
                
                If OpenValue = 0 And CloseValue <> 0 Then
                    Worksheets(a).Cells(c, 11).Value = 100
                ElseIf OpenValue <> 0 And CloseValue = 0 Then
                    Worksheets(a).Cells(c, 11).Value = -100
                ElseIf OpenValue = 0 And CloseValue = 0 Then
                    Worksheets(a).Cells(c, 11).Value = 0
                Else
                    PercentChar = (CloseValue - OpenValue) / OpenValue
                    Worksheets(a).Cells(c, 11).Value = PercentChar
                    Worksheets(a).Cells(c, 11).NumberFormat = "##.##%"
                End If
    
                'Initialize values for new ticker
                Ticker = Worksheets(a).Cells(b, 1).Value
                TotalSV = Worksheets(a).Cells(b, 7).Value
                OpenValue = Worksheets(a).Cells(b, 3).Value
            Else
                TotalSV = TotalSV + Worksheets(a).Cells(b, 7).Value
                Ticker = Worksheets(a).Cells(b, 1).Value
                CloseValue = Worksheets(a).Cells(b, 6).Value
     
            End If
        Next b
        
        ' Finding Greatest Values
        lastrow2 = Worksheets(a).Cells(Rows.Count, 10).End(xlUp).Row + 1
        For d = 2 To lastrow2
            GreatAux = Worksheets(a).Cells(d, 11).Value
            StockVolAux = Worksheets(a).Cells(d, 12).Value
            
            If GreatAux > GreatInc Then
                GreatInc = GreatAux
                GreatIncTicker = Worksheets(a).Cells(d, 9).Value
               
            ElseIf GreatAux < GreatDec Then
                GreatDec = GreatAux
                GreatDecTicker = Worksheets(a).Cells(d, 9).Value
            End If
            
            If StockVolAux > StockVol Then
                StockVol = StockVolAux
                StockVolTicker = Worksheets(a).Cells(d, 9).Value
            End If
            
        Next d
        
        Worksheets(a).Range("P2").Value = GreatInc
        Worksheets(a).Range("P2").NumberFormat = "##.##%"
        Worksheets(a).Range("P3").Value = GreatDec
        Worksheets(a).Range("P3").NumberFormat = "##.##%"
        Worksheets(a).Range("P4").Value = StockVol
        Worksheets(a).Range("O2").Value = GreatIncTicker
        Worksheets(a).Range("O3").Value = GreatDecTicker
        Worksheets(a).Range("O4").Value = StockVolTicker
        
Next a
    
End Sub
