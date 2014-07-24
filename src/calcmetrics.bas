' ===============================================
' Author: Yuanjia Su <yuanjia.su.13 at ucl.ac.uk>
'
' Description:
'         Retrive tick data and calculate metrics
' ===============================================

 Sub calcmetrics(i As Long, maxI As Long, retry As Integer)
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
            Application.OnTime Now + TimeValue("00:00:5"), _
            "'metrics " & i & ", " & maxI & ", " & retry & "'"
        End With
    End With
    
    
End Sub

Sub metrics(i As Long, maxI As Long, retry As Integer)
    Debug.Print "metrics: " & i & ", retry " & retry

    On Error Resume Next
    
    date3 = Format(CDate(Worksheets("Sheet3").Range("A2").Value), "dd/mm/yyyy")
    date2 = Format(CDate(Worksheets("Sheet2").Range("G" & i & "").Value), "dd/mm/yyyy")
    
    On Error GoTo 0
    
    ' match date and retry 3 times
    If Not date3 = date2 Then
        Debug.Print "Error: " & i & ", retry " & retry
        If retry < 3 Then
            Call calcmetrics(i, maxI, retry + 1)
        Else
            Call calcmetrics(i + 1, maxI, 1)
        End If
        Exit Sub
    End If
    
    ' delete dates out of trading range
    deltDates
    
    ' calculate
    ActiveSheet.EnableCalculation = False
        
    spreads (i)
    quotes (i)
    trades (i)

    ActiveSheet.EnableCalculation = True
    
    ' save every 20 calculations, roughly 5 mins
    If i Mod 20 = 0 Then
        Debug.Print "save...." & i
        ActiveWorkbook.Save
    End If
    
    If i + 1 <= maxI Then
        Call calcmetrics(i + 1, maxI, 1)
    Else
        Debug.Print "Finished. " & Now
        Debug.Print "saving... "
        ActiveWorkbook.Save
    End If
    
End Sub

Sub deltDates()
    
    With Worksheets("Sheet3")
    lastRow = .Cells(.Rows.count, "A").End(xlUp).Row
    
    For j = 2 To lastRow
        timestmp = CDate(Format(Range("A" & j & "").Value, "hh:mm:ss"))
        If timestmp > "16:50:00" Then
            Range("A" & j & ": F" & j & "").Formula = ""
        Else
            Exit For
        End If
    Next

    End With

End Sub


Sub errorCheck()
    
    With Worksheets("Sheet3")
    lastRow = .Cells(.Rows.count, "A").End(xlUp).Row
    
    ' delet empty rows
    Cells.Replace "#N/A", "", xlWhole
    
    End With
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


Sub formulas()

With Worksheets("Sheet3")

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet3")
    
    lastRow = .Cells(.Rows.count, "A").End(xlUp).Row
    lastTradeRow = .Cells(.Rows.count, "H").End(xlUp).Row

    ' -------------------------------------------------------
    ' BA spread
    ws.Range("J1").Formula = "BA Spread"
    ws.Range("J3").Formula = "=(E2-B2)/B2*10000"
    ws.Range("J3").Copy Range("J4:J" & lastRow & "")

    ' Bid diff
    ws.Range("K1").Formula = "Bid diff"
    ws.Range("K3").Formula = "=A2-A3"
    ws.Range("K3").Copy Range("K4:K" & lastRow & "")
    
    ' Ask diff
    ws.Range("L1").Formula = "Ask diff"
    ws.Range("L3").Formula = "=D2-D3"
    ws.Range("L3").Copy Range("L4:L" & lastRow & "")
    
    
    ' Trade diff
    ws.Range("M1").Formula = "Trade diff"
    ws.Range("M3").Formula = "=G2-G3"
    ws.Range("M3").Copy Range("M4:M" & lastTradeRow & "")
    
    ' -------------------------------------------------------
    ' BA spread
    ' -------------------------------------------------------
    
    ' /* Geo mean */
    
    ' valid spread
    ws.Range("N1").Formula = "spread valid"
    ws.Range("N2").Formula = "=IF(AND(ISNUMBER(J2),J2>0),1,0)"
    ws.Range("N2").Copy Range("N3:N" & lastRow & "")
    ' log spread
    ws.Range("O1").Formula = "log spd"
    ws.Range("O2").Formula = "=IF(N2=1,LOG(J2),0)"
    ws.Range("O2").Copy Range("O3:O" & lastRow & "")
    ' sum the log form of spreads
    ws.Range("P1").Formula = "spread log sum"
    ws.Range("P2").Formula = "=SUMIF(O:O,"">0"")"
    ws.Range("P3").Formula = "spread count"
    ws.Range("P4").Formula = "=COUNTIF(N:N,1)"
    ws.Range("P5").Formula = "geo spread"
    ws.Range("P6").Formula = "=EXP(P2/P4)"
    
    ' /* Avg mean */
    
    ws.Range("P7").Formula = "sum"
    ws.Range("P8").Formula = "=SUMIF(J:J,"">0"")"
    ws.Range("P9").Formula = "avg spread"
    ws.Range("P10").Formula = "=P8/P4"
    
    
    ' /* Duration wtd spread */
    ws.Range("T1").Formula = "spread wtd"
    ws.Range("T2").Formula = "=IF(K2=0,1,K2*100000+1)" 'convert date to number
    ws.Range("T2").Copy Range("T3:T" & lastRow & "")
    
    ws.Range("U1").Formula = "spread wtd sum"
    ws.Range("U2").Formula = "=T2*N2*J2"
    ws.Range("U2").Copy Range("U3:U" & lastRow & "")
    
    ws.Range("P23").Formula = "sum spread"
    ws.Range("P24").Formula = "=SUMIF(U:U,"">0"")"
    
    ws.Range("V1").Formula = "wtd sum"
    ws.Range("V2").Formula = "=T2*N2"
    ws.Range("V2").Copy Range("V3:V" & lastRow & "")
    ws.Range("P25").Formula = "sum weight"
    ws.Range("P26").Formula = "=SUMIF(V:V,"">0"")"
    
    ws.Range("P27").Formula = "duration wtd spread"
    ws.Range("P28").Formula = "=P24/P26"

    ' -------------------------------------------------------
    ' Quotes
    ' -------------------------------------------------------
    
    ' /* Number of Quotes */
    ws.Range("P11").Formula = "num of quots"
    ws.Range("P12").Formula = "=COUNT(B:B)"
    
    ' /* Avg Quote size */
    ws.Range("P13").Formula = "avg quote size"
    ws.Range("P14").Formula = "=(AVERAGE(C:C)+AVERAGE(F:F))/2"
    ws.Range("S1").Formula = "avg quote size"
    ws.Range("S2").Formula = "=IF(C2>0, (C2+F2)/2,0)"
    ws.Range("S2").Copy Range("S3:S" & lastRow & "")
    
    ' /* Geo Quote size */
    ws.Range("Q1").Formula = "log quots"
    ws.Range("Q2").Formula = "=IF(C2>0, LOG(S2),0)"
    ws.Range("Q2").Copy Range("Q3:Q" & lastRow & "")
    ws.Range("P15").Formula = "geo quote size"
    ws.Range("P16").Formula = "=EXP(SUM(Q:Q)/P4)"
    
    ' /* Duration wtd quotes*/
    ws.Range("W1").Formula = "quote size valid"
    ws.Range("W2").Formula = "=IF(AND(S2>0,T2>0),1,0)"
    ws.Range("W2").Copy Range("W3:W" & lastRow & "")
    ws.Range("X1").Formula = "quote size wtd sum"
    ws.Range("X2").Formula = "=S2*T2*W2"
    ws.Range("X2").Copy Range("X3:X" & lastRow & "")
    
    ws.Range("P29").Formula = "weight quote size sum"
    ws.Range("P30").Formula = "=SUMIF(X:X,"">0"")"
    ws.Range("P31").Formula = "wtd quote size"
    ws.Range("P32").Formula = "=P30/P26"
    
    ' -------------------------------------------------------
    ' Trades
    ' -------------------------------------------------------
    
    ' /* Number of Trades */
    ws.Range("P17").Formula = "num trades"
    ws.Range("P18").Formula = "=COUNT(I:I)"
    
    ' /* Avg of Trades */
    ws.Range("P19").Formula = "avg trade size"
    ws.Range("P20").Formula = "=AVERAGEIF(I:I,"">0"")"
    
    ' /* Geo of Trades */
    ws.Range("R1").Formula = "log trades"
    ws.Range("R2").Formula = "=IF(I2>0, LOG(I2),0)"
    ws.Range("R2").Copy Range("R3:R" & lastTradeRow & "")
    ws.Range("P21").Formula = "Geo trades"
    ws.Range("P22").Formula = "=EXP(SUMIF(R:R,"">0"")/P18)"
    
    ' /* Duration wtd trade size */
    ws.Range("Y1").Formula = "trade wtd"
    ws.Range("Y2").Formula = "=IF(M2=0,1,M2*100000+1)"
    ws.Range("Y2").Copy Range("Y3:Y" & lastRow & "")
    
    ws.Range("Z1").Formula = "trade wtd sum"
    ws.Range("Z2").Formula = "=I2*Y2"
    ws.Range("Z2").Copy Range("Z3:Z" & lastRow & "")
    
    ws.Range("AA1").Formula = "valid trade wtd sum"
    ws.Range("AA2").Formula = "=IF(I2>0,Y2, 0)"
    ws.Range("AA2").Copy Range("AA3:AA" & lastRow & "")
    
    ws.Range("P33").Formula = "wtd trade size sum"
    ws.Range("P34").Formula = "=SUMIF(Z:Z,"">0"")"
    ws.Range("P35").Formula = "trade wt sum"
    ws.Range("P36").Formula = "=SUMIF(AA:AA,"">0"")"
    ws.Range("P37").Formula = "wtd trade size"
    ws.Range("P38").Formula = "=P34/P36"

End With
End Sub


' algorithm mean, and median for spread
' geomatrical mean is not calculated because negative and zero numbers are included
Sub spreads(i As Long)
        
    With Worksheets("Sheet3")
        geo = Range("P6").Value
        avg = Range("P10").Value
        wtd = Range("P28").Value
        med = ActiveSheet.Evaluate("MEDIAN(IFERROR(IF(($J:$J <> """")*($J:$J > 0),$J:$J), """"))")
    End With
    
    Worksheets("Sheet2").Range("O" & i & "").Value = geo
    Worksheets("Sheet2").Range("M" & i & "").Value = avg
    Worksheets("Sheet2").Range("N" & i & "").Value = med
    Worksheets("Sheet2").Range("P" & i & "").Value = wtd
End Sub


Sub quotes(i As Long)

With Worksheets("Sheet3")
    num = Range("P12").Value
    avg = Range("P14").Value
    geo = Range("P16").Value
    wtd = Range("P32").Value
    med = ActiveSheet.Evaluate("MEDIAN(IFERROR(IF(($S:$S <> """")*($S:$S > 0),$S:$S), """"))")
End With

Worksheets("Sheet2").Range("Q" & i & "").Value = num
Worksheets("Sheet2").Range("T" & i & "").Value = geo
Worksheets("Sheet2").Range("R" & i & "").Value = avg
Worksheets("Sheet2").Range("S" & i & "").Value = med
Worksheets("Sheet2").Range("U" & i & "").Value = wtd
End Sub


Sub trades(i As Long)
 
    With Worksheets("Sheet3")
        num = Range("P18").Value
        avg = Range("P20").Value
        geo = Range("P22").Value
        wtd = Range("P38").Value
        med = ActiveSheet.Evaluate("MEDIAN(IFERROR(IF(($I:$I <> """")*($I:$I > 0),$I:$I), """"))")

    End With
    
    Worksheets("Sheet2").Range("V" & i & "").Value = num
    Worksheets("Sheet2").Range("W" & i & "").Value = avg
    Worksheets("Sheet2").Range("Y" & i & "").Value = geo
    Worksheets("Sheet2").Range("X" & i & "").Value = med
    Worksheets("Sheet2").Range("Z" & i & "").Value = wtd
    
End Sub

