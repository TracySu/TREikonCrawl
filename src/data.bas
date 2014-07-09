' ##########################################################
' Author: Yuanjia Su
' Contact: yuanjia.su.13 at ucl.ac.uk
'
' --
' Copyright (c) 2014 Yuanjia Su
' ##########################################################


' retrieve indecies with constituent names
Sub getIndexNames()
    With Worksheets("Sheet1")
    
    .Range("A1").Formula = _
    "=TR("".FTSE"",""TR.IndexConstituentRIC"",""RH=In"")"
    
    .Range("A102").Formula = _
    "=TR("".SSMI"",""TR.IndexConstituentRIC"",""RH=In"")"
    
    .Range("A122").Formula = _
    "=TR("".GDAXI"",""TR.IndexConstituentRIC"",""RH=In"")"
    
    .Range("A153").Formula = _
    "=TR("".FCHI"",""TR.IndexConstituentRIC"",""RH=In"")"
    
    .Range("A193").Formula = _
    "=TR("".IBEX"",""TR.IndexConstituentRIC"",""RH=In"")"
    
    .Range("A228").Formula = _
    "=TR("".OMXSPI"",""TR.IndexConstituentRIC"",""RH=In"")"

    .Range("A516").Formula = _
    "=TR("".OMXHPI"",""TR.IndexConstituentRIC"",""RH=In"")"

    .Range("A646").Formula = _
    "=TR("".PSI20"",""TR.IndexConstituentRIC"",""RH=In"")"
    
    .Range("A666").Formula = _
    "=TR("".ATX"",""TR.IndexConstituentRIC"",""RH=In"")"
    
    .Range("A686").Formula = _
    "=TR("".ATG"",""TR.IndexConstituentRIC"",""RH=In"")"
    
    .Range("A746").Formula = _
    "=TR("".XU100"",""TR.IndexConstituentRIC"",""RH=In"")"
    
    .Range("A846").Formula = _
    "=TR("".WIG"",""TR.IndexConstituentRIC"",""RH=In"")"
    
    .Range("A1224").Formula = _
    "=TR("".BFX"",""TR.IndexConstituentRIC"",""RH=In"")"
    
    .Range("A1244").Formula = _
    "=TR("".PX"",""TR.IndexConstituentRIC"",""RH=In"")"

'    .Range("A1257").Formula = _
'    "=TR("".FTMIB"",""TR.IndexConstituentRIC"",""RH=In"")"

    End With

End Sub


'
Sub checkZeros()

With Worksheets("Sheet1")
    Dim lastRow As Integer
    lastRow = .Cells(.Rows.count, "A").End(xlUp).Row
    
    For i = 1 To lastRow
        
        If (Range("A" & i & "").Value = 0) Then
            Range("A" & i & "").Value = Range("A" & i + 1 & "").Value
        Else
        End If
    Next
    
End With
End Sub


' anto fill dates that are not retrived in cells
Sub dateCheck()
With Worksheets("Sheet2")

    For i = 1 To 75291
    If Range(" G" & i & " ").Value = "00:00:00" Then
        Range("G" & i + 1 & "").Select
        Selection.AutoFill Destination:=Range("G" & i & ":G" & i + 1 & ""), Type:=xlFillDefault
    End If
    Next
End With
End Sub


' fill currency in case data missed
Sub currCheck()

With Worksheets("Sheet2")
    For i = 1 To 75291

    If Range("E" & i & "").Formula = "United Kingdom" Then
        Range("C" & i & " ").Formula = "GBp"
    ElseIf Range("E" & i & "").Formula = "Switzerland" Then
        Range(" C" & i & " ").Value = "CHF"
    ElseIf Range("E" & i & "").Formula = "Germany" Then
        Range(" C" & i & " ").Value = "EUR"
    ElseIf Range("E" & i & "").Formula = "France" Then
        Range(" C" & i & " ").Value = "EUR"
    ElseIf Range("E" & i & "").Formula = "Netherlands" Then
        Range(" C" & i & " ").Value = "EUR"
    ElseIf Range("E" & i & "").Formula = "Belgium" Then
        Range(" C" & i & " ").Value = "EUR"
    ElseIf Range("E" & i & "").Formula = "Spain" Then
        Range(" C" & i & " ").Value = "EUR"
    ElseIf Range("E" & i & "").Formula = "Sweden" Then
        Range(" C" & i & " ").Value = "SEK"
    ElseIf Range("E" & i & "").Formula = "Finland" Then
        Range(" C" & i & " ").Value = "EUR"
    ElseIf Range("E" & i & "").Formula = "Portugal" Then
        Range(" C" & i & " ").Value = "EUR"
    ElseIf Range("E" & i & "").Formula = "Turkey" Then
        Range(" C" & i & " ").Value = "TRY"
    ElseIf Range("E" & i & "").Formula = "Poland" Then
        Range(" C" & i & " ").Value = "PLN"
    ElseIf Range("E" & i & "").Formula = "Greece" Then
        Range(" C" & i & " ").Value = "EUR"
    ElseIf Range("E" & i & "").Formula = "Czech Republic" Then
        Range(" C" & i & " ").Value = "CZK"
    End If
    
    Next
End With

End Sub

' function to return to array of stock names
Function getStockNames() As String()

    Dim arr(1256) As String
    Dim i As Integer
    i = 0

    For Each c In Worksheets("Sheet1").Range("B2:B1256")
        arr(i) = Chr(34) & c.Value & Chr(34)
        i = i + 1
    Next c
    
    getStockNames = arr
    
End Function

' function to return all index names
Function getIndNames() As String()

    Dim arr(1256) As String
    Dim i As Integer
    i = 0

    For Each c In Worksheets("Sheet1").Range("A2:A1256")
        arr(i) = Chr(34) & c.Value & Chr(34)
        i = i + 1
    Next c
    
    getIndNames = arr
    
End Function


' retrieve intraday data
Sub getStockPrices()
    
    stocks = getStockNames
    indxs = getIndNames
    numStocks = Application.CountA(stocks) - 1
    flag = 2
    
'    For i = 0 To numStocks - 1
'        ' Stock symbols
'        temp = flag + 60
'        Worksheets("Sheet2").Range("A" & flag & ":A" & temp & " ").Formula = _
'        " " & stocks(i) & " "
'
'        ' Index names
'        Worksheets("Sheet2").Range("B" & flag & ":B" & temp & " ").Formula = _
'        " " & indxs(i) & " "
'
'         ' get stock info
'         Worksheets("Sheet2").Range("C" & flag & ":C" & temp & " ").Formula = _
'         "=TR(" & stocks(i) & ",""CURRENCY;TR.CompanyMarketCap;TR.ExchangeCountry"")"
'
'         ' get stock prices
'         Worksheets("Sheet2").Range("G" & flag & " ").Formula = _
'         "=RHistory(" & stocks(i) & ",""TRDPRC_1.Timestamp;TRDPRC_1.Open;TRDPRC_1.High;TRDPRC_1.Low;TRDPRC_1.Close;TRDPRC_1.Volume"",""START:02-Apr-2014 END:30-Jun-2014 INTERVAL:1D"",,""SORT:ASC"")"
'
'         flag = flag + 60
'         ActiveWorkbook.Save
'    Next i
    
    ' ----------------
    ' Set colomn names
    ' ----------------
    Worksheets("Sheet2").Range("A1").Formula = _
    "Stock"
    
    Worksheets("Sheet2").Range("B1").Formula = _
    "Index"
    
    Worksheets("Sheet2").Range("C1").Formula = _
    "Currency"
    
    Worksheets("Sheet2").Range("D1").Formula = _
    "MarktCap"
    
    Worksheets("Sheet2").Range("E1").Formula = _
    "ExchangeCountry"
    
    Worksheets("Sheet2").Range("F1").Formula = _
    "CAP"
    
    Worksheets("Sheet2").Range("G1").Formula = _
    "Timestamp"
    
    Worksheets("Sheet2").Range("H1").Formula = _
    "Open"
    
    Worksheets("Sheet2").Range("I1").Formula = _
    "High"
    
    Worksheets("Sheet2").Range("J1").Formula = _
    "Low"
    
    Worksheets("Sheet2").Range("K1").Formula = _
    "Close"
    
    Worksheets("Sheet2").Range("L1").Formula = _
    "Volume"
    
    Worksheets("Sheet2").Range("M1").Formula = _
    "AvgBASpread(BP)"
    
    Worksheets("Sheet2").Range("N1").Formula = _
    "GeoAvgBASpread(BP)"
    
    Worksheets("Sheet2").Range("O1").Formula = _
    "NumberOfTrades"
    
    Worksheets("Sheet2").Range("P1").Formula = _
    "NumberOfQuotes"
    
    Worksheets("Sheet2").Range("Q1").Formula = _
    "AvgBidVolume"
    
    Worksheets("Sheet2").Range("R1").Formula = _
    "AvgAskVolume"
    
    Worksheets("Sheet2").Range("S1").Formula = _
    "AvgBidPrice"
    
    Worksheets("Sheet2").Range("T1").Formula = _
    "AvgAskPrice"
    
    Worksheets("Sheet2").Range("U1").Formula = _
    "AvgTradePrice"
    
    Worksheets("Sheet2").Range("V1").Formula = _
    "AvgTradeVolume"
    
    Worksheets("Sheet2").Range("W1").Formula = _
    "AvgVWAP"
    
    Worksheets("Sheet2").Range("X1").Formula = _
    "AvgBidTimeDiff"
    
    Worksheets("Sheet2").Range("Y1").Formula = _
    "AvgAskTimeDiff"
    
    Worksheets("Sheet2").Range("Z1").Formula = _
    "AvgTradeTimeDiff"
    
    Worksheets("Sheet2").Range("AA1").Formula = _
    "WtdBidPrice"
    
    Worksheets("Sheet2").Range("AB1").Formula = _
    "WtdAskPrice"
    
    Worksheets("Sheet2").Range("AC1").Formula = _
    "WtdB/ASpread"
    
    Worksheets("Sheet2").Range("AD1").Formula = _
    "MedianBASpread"
    
    Worksheets("Sheet2").Range("AE1").Formula = _
    "MedianBidPrice"
    
    Worksheets("Sheet2").Range("AF1").Formula = _
    "MedianAskPrice"
    
    Worksheets("Sheet2").Range("AG1").Formula = _
    "MedTradePrice"
    
    Worksheets("Sheet2").Range("AH1").Formula = _
    "GeoBidPrice"

    Worksheets("Sheet2").Range("AI1").Formula = _
    "GeoAskPrice"
    
    Worksheets("Sheet2").Range("AJ1").Formula = _
    "GeoTradePrice"
    
End Sub


' capitalisation of stocks
Sub getCAP()

    With Worksheets("Sheet2")
        lastRow = .Cells(.Rows.count, "C").End(xlUp).Row

        Dim currcy, capType As String
        Dim marketCap, cap, ratio
        
        For i = 2 To lastRow
            
            currcy = .Range("C" & i & "").Value
            marketCap = .Range("D" & i & "").Value
              
            Select Case currcy
            Case "GBp"
                ratio = Worksheets("Sheet1").Range("F2").Value
                cap = ratio * marketCap
            Case "Euro"
                ratio = Worksheets("Sheet1").Range("F3").Value
                cap = ratio * marketCap
            Case Else
            
            End Select
        
            If cap < 50000000 Then
                capType = "Nano-cap"
                
            ElseIf cap > 50000000 And cap < 250000000 Then
                capType = "Micro-cap"
            
            ElseIf cap > 250000000 And cap < 2000000000 Then
                capType = "Small-cap"
            
            ElseIf cap > 2000000000 And cap < 10000000000# Then
                capType = "Mid-cap"
                        
            ElseIf cap > 10000000000# And cap < 200000000000# Then
                capType = "Large-cap"
            
            ElseIf cap > 200000000000# Then
                capType = "Mega-cap"
            
            End If
            
            Worksheets("Sheet2").Range("F" & i & "").Value = capType

    Next
    End With
    
End Sub


' delet dates that Eikon can no longer retrived
' as the past 3 month only restriction
Sub deletDate()

With Worksheets("Sheet2")

 For i = 55 To 69000
    If CDate(Range("G" & i & "").Value) < "15/04/2014" Then
    Range("G" & i & "").EntireRow.Delete
    
    If Range("G" & i & "").Value = "" Or CDate(Range("G" & i & "").Value) < "15/04/2014" Then
    Range("G" & i & "").EntireRow.Delete
    End If
  Next
  
End With
End Sub
