' ##########################################################
' Author: Yuanjia Su
' Contact: yuanjia.su.13 at ucl.ac.uk
'
' --
' Copyright (c) 2014 Yuanjia Su
' ##########################################################


Sub calcMetrics(i As Integer, maxI As Integer, retry As Integer)
    Debug.Print "calcMetrics: " & i & ", retry " & retry
    
    Dim name As String
    Dim tickDate As String
    
    With Worksheets("Sheet2")
    
        ' get stock name and date from sheet2
        name = .Range("A" & i & "").Value
        tickDate = Format(.Range("G" & i & "").Value, "dd-mmm-yyyy")
                    
        With Worksheets("Sheet3")
            ' retrieve tick data
            Call getTickData(name, tickDate)
            
            ' wait 4 seconds until data retrieving finished
            ' then calculate metrics
            Application.OnTime Now + TimeValue("00:00:4"), _
            "'metrics " & i & ", " & maxI & ", " & retry & "'"
        End With
    End With
    
    
End Sub

Sub metrics(i As Integer, maxI As Integer, retry As Integer)
    Debug.Print "metrics: " & i & ", retry " & retry

    On Error Resume Next
    
    date3 = Format(CDate(Worksheets("Sheet3").Range("A2").Value), "dd/mm/yyyy")
    date2 = Format(CDate(Worksheets("Sheet2").Range("G" & i & "").Value), "dd/mm/yyyy")
    
    On Error GoTo 0
    
    ' match date and retry 3 times
    If Not date3 = date2 Then
        Debug.Print "Error: " & i & ", retry " & retry
        If retry < 3 Then
            Call calcMetrics(i, maxI, retry + 1)
        Else
            Call calcMetrics(i + 1, maxI, 1)
        End If
        Exit Sub
    End If
        
    ' metrics
    spread (i)
    getVWAP
    avgSpread (i)
    numTrades (i)
    numQuotes (i)
    avgBidSize (i)
    avgAskSize (i)
    avgBidPrice (i)
    avgAskPrice (i)
    avgTradePrice (i)
    avgTradeVolume (i)
    avgVWAP (i)
    avgTimeDiff (i)
    weitdSpread (i)
    
    
    ' save the file every 20 calculations, roughly every 5 mins
    If i Mod 20 = 0 Then
        Debug.Print "save...." & i
        ActiveWorkbook.Save
    End If
    
    If i + 1 <= maxI Then
        Call calcMetrics(i + 1, maxI, 1)
    Else
        Debug.Print "Finished. " & Now
        Debug.Print "saving... "
        ActiveWorkbook.Save
    End If
    
End Sub


Sub getTickData(n As String, d As String)
    Debug.Print "getTickData: "
    
    With Worksheets("Sheet3")
    
    ' clear cells before retrive new data
    .Range("A:A").Formula = ""
    .Range("D:D").Formula = ""
    .Range("G:G").Formula = ""
    
    ' Bid price/volume
    .Range("A1").Formula = _
    "=RHistory(" & n & ",""BID.Timestamp;BID.Value;BID.Volume"",""TIMEZONE:LOCAL START:" & d & " END:" & d & " INTERVAL:TICK"",,""CH:Fd"")"

    ' Ask price/volume
    .Range("D1").Formula = _
    "=RHistory(" & n & ",""ASK.Timestamp;ASK.Value;ASK.Volume"",""TIMEZONE:LOCAL START:" & d & " END:" & d & " INTERVAL:TICK"",,""CH:Fd"")"
    
    ' Trade price/volume
    .Range("G1").Formula = _
    "=RHistory(" & n & ",""TRDPRC_1.Timestamp;TRDPRC_1.Value;TRDPRC_1.Volume"",""TIMEZONE:LOCAL START:" & d & " END:" & d & " INTERVAL:TICK"",,""CH:Fd"")"
    
    End With
    
End Sub

Sub spread(i As Integer)

    Dim lastRow
    With Worksheets("Sheet3")
        
        bidMaxRow = .Cells(.Rows.count, "B").End(xlUp).Row
        askMaxRow = .Cells(.Rows.count, "E").End(xlUp).Row
        .Range("J:J").Formula = ""
        
        If bidMaxRow > askMaxRow Then
            lastRow = askMaxRow
        Else
            lastRow = bidMaxRow
        End If
        
        ' calculate spread in base point
        .Range("J2:J" & lastRow & " ").Formula = _
        "= (E:E-B:B)/E:E * 10000"
        
        Range("J1").Value = "Bid/Ask Spread"
    End With

End Sub

Sub getVWAP()
    
    With Worksheets("Sheet3")
        lastPriceRow = .Cells(.Rows.count, "I").End(xlUp).Row
        .Range("K:K").Formula = ""
        .Range("L:L").Formula = ""
        .Range("M:M").Formula = ""
        .Range("N:N").Formula = ""
        
        For j = 2 To lastPriceRow
            
            If Not IsNumeric(Range("H" & j & "").Formula) Then
                Range("H" & j & "").Formula = ""
            End If
        
            If Not IsNumeric(Range("I" & j & "").Formula) Then
                Range("I" & j & "").Formula = ""
            End If
            
            ' VAMP calculation
            .Range("K" & j & " ").Formula = _
            "= H" & j & "*I" & j & ""

            .Range("L" & j & " ").Formula = _
            "=SUM(K2:K" & j & ")"
            
            .Range("M" & j & " ").Formula = _
            "=SUM(I2:I" & j & ")"
                      
            .Range("N" & j & " ").Value = _
            "=(L" & j & ")/(M" & j & ")"

        Next
        
        Range("K1").Value = "V*P"
        Range("L1").Value = "totalVP"
        Range("M1").Value = "totalV"
        Range("N1").Value = "VWAP"
    End With

End Sub

' algorithm mean, and median for spread
' geomatrical mean is not calculated because negative and zero numbers are included
Sub avgSpread(i As Integer)
    
    With Worksheets("Sheet3")
        spd = ActiveSheet.Evaluate("AVERAGE(J:J)")
        med = ActiveSheet.Evaluate("MEDIAN(J:J)")
    End With
    
    Worksheets("Sheet2").Range("M" & i & "").Value = spd
    Worksheets("Sheet2").Range("AD" & i & "").Value = med
    
End Sub


Sub numTrades(i As Integer)
    
    Dim lastTradeRow
    
    With Worksheets("Sheet3")
        lastTradeRow = .Cells(.Rows.count, "H").End(xlUp).Row
    End With
    
    Worksheets("Sheet2").Range("O" & i & "").Value = lastTradeRow

End Sub


Sub numQuotes(i As Integer)
    
    Dim lastTradeRow
    
    With Worksheets("Sheet3")
        lastQuoteRow = .Cells(.Rows.count, "B").End(xlUp).Row
    End With
    
    Worksheets("Sheet2").Range("P" & i & "").Value = lastQuoteRow

End Sub

Sub avgBidSize(i As Integer)
    
    With Worksheets("Sheet3")
        avgSize = ActiveSheet.Evaluate("AVERAGE(C:C)")
    End With

    Worksheets("Sheet2").Range("Q" & i & "").Value = avgSize
End Sub

Sub avgAskSize(i As Integer)
    
    With Worksheets("Sheet3")
        avgSize = ActiveSheet.Evaluate("AVERAGE(F:F)")
    End With

    Worksheets("Sheet2").Range("R" & i & "").Value = avgSize
End Sub

Sub avgBidPrice(i As Integer)
    
    With Worksheets("Sheet3")
        avgPrice = ActiveSheet.Evaluate("AVERAGE(B:B)")
        medPrice = ActiveSheet.Evaluate("MEDIAN(B:B)")
        geoPrice = ActiveSheet.Evaluate("PRODUCT(B2:B3000^(1/COUNT(B2:B3000)))")
        'geoPrice = ActiveSheet.Evaluate("GEOMEAN(B:B)")
    End With

    Worksheets("Sheet2").Range("S" & i & "").Value = avgPrice
    Worksheets("Sheet2").Range("AE" & i & "").Value = medPrice
    Worksheets("Sheet2").Range("AH" & i & "").Value = geoPrice
    
End Sub

Sub avgAskPrice(i As Integer)
    
    With Worksheets("Sheet3")
        avgPrice = ActiveSheet.Evaluate("AVERAGE(E:E)")
        medPrice = ActiveSheet.Evaluate("MEDIAN(E:E)")
        geoPrice = ActiveSheet.Evaluate("PRODUCT(E2:E3000^(1/COUNT(E2:E3000)))")
    End With

    Worksheets("Sheet2").Range("T" & i & "").Value = avgPrice
    Worksheets("Sheet2").Range("AF" & i & "").Value = medPrice
    Worksheets("Sheet2").Range("AI" & i & "").Value = geoPrice
    
End Sub

Sub avgTradePrice(i As Integer)
    
    With Worksheets("Sheet3")
    lastRow = .Cells(.Rows.count, "H").End(xlUp).Row - 1
        avgPrice = ActiveSheet.Evaluate("AVERAGE(H:H)")
        medPrice = ActiveSheet.Evaluate("MEDIAN(H:H)")
        geoPrice = ActiveSheet.Evaluate("PRODUCT(H2:H" & lastRow & "^(1/COUNT(H2:H" & lastRow & ")))")
    End With

    Worksheets("Sheet2").Range("U" & i & "").Value = avgPrice
    Worksheets("Sheet2").Range("AG" & i & "").Value = medPrice
    Worksheets("Sheet2").Range("AJ" & i & "").Value = geoPrice
    
End Sub
Sub avgTradeVolume(i As Integer)
    
    With Worksheets("Sheet3")
        avgSize = ActiveSheet.Evaluate("AVERAGE(I:I)")
    End With

    Worksheets("Sheet2").Range("V" & i & "").Value = avgSize
End Sub

'[change]
Sub avgVWAP(i As Integer)

    With Worksheets("Sheet3")
    lastRow = .Cells(.Rows.count, "N").End(xlUp).Row
    For j = 2 To lastRow
        If Not IsNumeric(Range("N" & j & "").Value) Then
            Range("N" & j & "").Formula = ""
        End If
    Next
    
    avgPrice = ActiveSheet.Evaluate("AVERAGE(N2:N" & lastRow & ")")
    End With

    Worksheets("Sheet2").Range("W" & i & "").Value = avgPrice
End Sub


Sub avgTimeDiff(i As Integer)

    Dim avgBid
    Dim avgAsk
    Dim avgTrade
    
    With Worksheets("Sheet3")
        .Range("O:O").Formula = ""
        .Range("P:P").Formula = ""
        .Range("Q:Q").Formula = ""
        
        lastBidRow = .Cells(.Rows.count, "B").End(xlUp).Row
        lastAskRow = .Cells(.Rows.count, "E").End(xlUp).Row
        lastTradeRow = .Cells(.Rows.count, "H").End(xlUp).Row
        
        For j = 3 To lastBidRow - 1
            .Range("O" & j & " ").Value = _
            "= A" & j & " - A" & j + 1 & ""
            
        Next
        
        avgBid = ActiveSheet.Evaluate("AVERAGE(O:O)")
        
        For j = 3 To lastBidRow - 1
            .Range("P" & j & " ").Value = _
            "= D" & j & " - D" & j + 1 & ""
        Next
        
        avgAsk = ActiveSheet.Evaluate("AVERAGE(P:P)")
        
        For j = 3 To lastBidRow - 1
            .Range("Q" & j & " ").Value = _
            "= G" & j & " - G" & j + 1 & ""
        Next
        
        avgTrade = ActiveSheet.Evaluate("AVERAGE(Q:Q)")
        
        Range("O1").Value = "Bid Time Diff"
        Range("P1").Value = "Ask Time Diff"
        Range("Q1").Value = "Trade Time Diff"
    End With
    
    Worksheets("Sheet2").Range("X" & i & "").Value = avgBid
    Worksheets("Sheet2").Range("Y" & i & "").Value = avgAsk
    Worksheets("Sheet2").Range("Z" & i & "").Value = avgTrade
    
End Sub


Sub weitdSpread(i As Integer)

    Dim count As Integer
    Dim first
    Dim last

    
    With Worksheets("Sheet3")
    .Range("R:R").Formula = ""
    .Range("S:S").Formula = ""
    .Range("T:T").Formula = ""
        
    lastBidRow = .Cells(.Rows.count, "A").End(xlUp).Row
        
    count = 1
    first = 2
    last = 2

    
    For j = 2 To lastBidRow
        prev = Range("A" & j - 1 & "").Value
        curr = Range("A" & j & "").Value
        
        If (curr = prev) Then
            count = count + 1
            last = j
        Else
            
            .Range("R" & first & " : R" & last & " ").Value = _
            "= B" & first & ":B" & last & " * (" & count & "/" & lastBidRow & ") "

            .Range("S" & first & " : S" & last & " ").Value = _
            "= E" & first & ":E" & last & " * (" & count & "/" & lastBidRow & ") "
            
            .Range("T" & first & " : T" & last & " ").Value = _
            "= J" & first & ":J" & last & " * (" & count & "/" & lastBidRow & ") "
        
        last = j
        first = j
        
        count = 1
        End If
    Next
    
    Range("R1").Value = "Bid*Duration"
    Range("S1").Value = "Ask*Duration"
    Range("T1").Value = "Spread*Duration"
    
    bid = ActiveSheet.Evaluate("SUM(R:R)")
    ask = ActiveSheet.Evaluate("SUM(S:S)")
    spd = ActiveSheet.Evaluate("SUM(T:T)")
    
    End With
    
    Worksheets("Sheet2").Range("AA" & i & "").Value = bid
    Worksheets("Sheet2").Range("AB" & i & "").Value = ask
    Worksheets("Sheet2").Range("AC" & i & "").Value = spd
       
End Sub

