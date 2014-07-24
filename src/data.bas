' ===============================================
' Author: Yuanjia Su <yuanjia.su.13 at ucl.ac.uk>
'
' Description:
'         1. Get indecies and symbols
'         2. Get daily data
' ===============================================


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

    .Range("A1252").Formula = _
    "=TR("".AEX"",""TR.IndexConstituentRIC"",""RH=In"")"

    .Range("A1277").Formula = _
    "=TR("".OMXC20"",""TR.IndexConstituentRIC"",""RH=In"")"
    
    .Range("A1297").Formula = _
    "=TR("".OBX"",""TR.IndexConstituentRIC"",""RH=In"")"

    .Range("A1322").Formula = _
    "=TR("".FTMC"",""TR.IndexConstituentRIC"",""RH=In"")"
    
    .Range("A1571").Formula = _
    "=TR("".SDAXI"",""TR.IndexConstituentRIC"",""RH=In"")"

    End With

End Sub


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

Sub delZeros()

With Worksheets("Sheet2")
    lastRow = .Cells(.Rows.count, "A").End(xlUp).Row
    For i = 68148 To 69193
        If (Range("G" & i & "").Value = 0 Or Range("B" & i & "").Value = 0) Then
        Range("A" & i & "").EntireRow.Delete
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

    If Range("C" & i & "").Formula = "GBp" Then
        Range("E" & i & " ").Formula = "United Kingdom"
        
    ElseIf Range("C" & i & "").Formula = "CHF" Then
        Range(" E" & i & " ").Value = "Switzerland"
        
    ElseIf Range("C" & i & "").Formula = "CZK" Then
        Range(" E" & i & " ").Value = "Czech Republic"
        
    ElseIf Range("C" & i & "").Formula = "DKK" Then
        Range(" E" & i & " ").Value = "Denmark"
        
    ElseIf Range("C" & i & "").Formula = "Netherlands" Then
        Range(" E" & i & " ").Value = "EUR"
        
    ElseIf Range("C" & i & "").Formula = "Belgium" Then
        Range(" E" & i & " ").Value = "EUR"
        
    ElseIf Range("C" & i & "").Formula = "Spain" Then
        Range(" E" & i & " ").Value = "EUR"
        
    ElseIf Range("C" & i & "").Formula = "Sweden" Then
        Range(" E" & i & " ").Value = "SEK"
        
    ElseIf Range("C" & i & "").Formula = "Finland" Then
        Range(" E" & i & " ").Value = "EUR"
        
    ElseIf Range("C" & i & "").Formula = "Portugal" Then
        Range(" E" & i & " ").Value = "EUR"
        
    ElseIf Range("C" & i & "").Formula = "Turkey" Then
        Range(" E" & i & " ").Value = "TRY"
        
    ElseIf Range("C" & i & "").Formula = "Poland" Then
        Range(" E" & i & " ").Value = "PLN"
        
    ElseIf Range("C" & i & "").Formula = "Greece" Then
        Range(" E" & i & " ").Value = "EUR"
        
    ElseIf Range("C" & i & "").Formula = "Czech Republic" Then
        Range(" E" & i & " ").Value = "CZK"
    End If
    
    Next
End With

End Sub

' array of stock names
Function getStockNames() As String()

    Dim arr(1615) As String
    Dim i As Integer
    i = 0

    For Each c In Worksheets("Sheet1").Range("B2:B1615")
        arr(i) = Chr(34) & c.Value & Chr(34)
        i = i + 1
    Next c
    
    getStockNames = arr
    
End Function

' array of index names
Function getIndNames() As String()

    Dim arr(1615) As String
    Dim i As Integer
    i = 0

    For Each c In Worksheets("Sheet1").Range("A2:A1615")
        arr(i) = Chr(34) & c.Value & Chr(34)
        i = i + 1
    Next c
    
    getIndNames = arr
    
End Function


' retrieve daily data
Sub getStockPrices()
    
    stocks = getStockNames
    indxs = getIndNames
    numStocks = Application.CountA(stocks) - 1
    flag = 67104

    For i = 1252 To numStocks - 1

    Debug.Print i
        ' Stock symbols
        temp = flag + 58
        Worksheets("Sheet2").Range("A" & flag & ":A" & temp & " ").Formula = _
        " " & stocks(i) & " "

        ' Index names
        Worksheets("Sheet2").Range("B" & flag & ":B" & temp & " ").Formula = _
        " " & indxs(i) & " "

         ' get stock info
         Worksheets("Sheet2").Range("C" & flag & ":C" & temp & " ").Formula = _
         "=TR(" & stocks(i) & ",""CURRENCY;TR.CompanyMarketCap;TR.ExchangeCountry"")"

         ' get stock prices
         Worksheets("Sheet2").Range("G" & flag & " ").Formula = _
         "=RHistory(" & stocks(i) & ",""TRDPRC_1.Timestamp;TRDPRC_1.Open;TRDPRC_1.High;TRDPRC_1.Low;TRDPRC_1.Close;TRDPRC_1.Volume"",""START:02-May-2014 END:30-Jul-2014 INTERVAL:1D"",,""SORT:ASC"")"

         flag = flag + 55
         ActiveWorkbook.Save
    Next i
    
    ' ------------------------------------------------
    ' Set colomn names
    ' ------------------------------------------------
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
    "AvgSpread(BPS)"
    
    Worksheets("Sheet2").Range("N1").Formula = _
    "MedSpread(BPS)"
    
    Worksheets("Sheet2").Range("O1").Formula = _
    "GeoSpread(BPS)"
    
    Worksheets("Sheet2").Range("P1").Formula = _
    "WtdSpread(BPS)"
    
    Worksheets("Sheet2").Range("Q1").Formula = _
    "NumOfQuotes"
    
    Worksheets("Sheet2").Range("R1").Formula = _
    "AvgQuoteSize"
    
    Worksheets("Sheet2").Range("S1").Formula = _
    "MedQuoteSize"
    
    Worksheets("Sheet2").Range("T1").Formula = _
    "GeoQuoteSize"
    
    Worksheets("Sheet2").Range("U1").Formula = _
    "WtdQuoteSize"
    
    Worksheets("Sheet2").Range("V1").Formula = _
    "NumOfTrades"
    
    Worksheets("Sheet2").Range("W1").Formula = _
    "AvgTradeSize"
    
    Worksheets("Sheet2").Range("X1").Formula = _
    "MedTradeSize"
    
    Worksheets("Sheet2").Range("Y1").Formula = _
    "GeoTradeSize"
    
    Worksheets("Sheet2").Range("Z1").Formula = _
    "WtdTradeSize"
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
                cap = ratio * marketCap * 100
            Case "CHF"
                ratio = Worksheets("Sheet1").Range("F3").Value
                cap = ratio * marketCap
            Case "EUR"
                ratio = Worksheets("Sheet1").Range("F4").Value
                cap = ratio * marketCap
            Case "SEK"
                ratio = Worksheets("Sheet1").Range("F5").Value
                cap = ratio * marketCap
            Case "TRY"
                ratio = Worksheets("Sheet1").Range("F6").Value
                cap = ratio * marketCap
            Case "PLN"
                ratio = Worksheets("Sheet1").Range("F7").Value
                cap = ratio * marketCap
            Case "CZK"
                ratio = Worksheets("Sheet1").Range("F8").Value
                cap = ratio * marketCap
            Case "NOK"
                ratio = Worksheets("Sheet1").Range("F9").Value
                cap = ratio * marketCap
            Case "DKK"
                ratio = Worksheets("Sheet1").Range("F10").Value
                cap = ratio * marketCap
            End Select
        
            If cap < 50000000 Then
                capType = "Nano"
                
            ElseIf cap > 50000000 And cap < 250000000 Then
                capType = "Micro"
            
            ElseIf cap > 250000000 And cap < 2000000000 Then
                capType = "Small"
            
            ElseIf cap > 2000000000 And cap < 10000000000# Then
                capType = "Mid"
                        
            ElseIf cap > 10000000000# And cap < 200000000000# Then
                capType = "Large"
            
            ElseIf cap > 200000000000# Then
                capType = "Mega"
            
            End If
            
            Worksheets("Sheet2").Range("F" & i & "").Value = capType

    Next
    End With
    
End Sub


' delet dates Eikon can no longer retrive (3-month restriction)
Sub deletDate()

With Worksheets("Sheet2")

 For i = 55413 To 55463
'    If CDate(Range("G" & i & "").Value) < "15/04/2014" Then
'    Range("G" & i & "").EntireRow.Delete
    Debug.Print i
    If Range("G" & i & "").Value = "" Or Range("H" & i & "").Formula = "#N/A" Or Range("D" & i & "").Value = "" Then
        Range("G" & i & "").EntireRow.Delete
    End If
  Next
  
End With
End Sub

' convert formulas to static values
Sub stable()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet2")
    ws.UsedRange.Value = ws.UsedRange.Value
  
End Sub

