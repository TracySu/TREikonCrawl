' ##########################################################
' Author: Yuanjia Su
' Contact: yuanjia.su.13 at ucl.ac.uk
'
' --
' Copyright (c) 2014 Yuanjia Su
' ##########################################################

Sub main()
    With Worksheets("Sheet2")

        Debug.Print "Start: " & Now
        Dim maxStockRow As Integer
        'maxDateRow = .Cells(.Rows.Count, "J").End(xlUp).Row - 1
        
        ' specify start row and end row from sheet2
        Call calcMetrics(15000, 15200, 1)
    End With
End Sub


' main for sheet2
Sub mainSheet2()

    Call getStockPrices
    startTime = Now + TimeValue("00:00:3")
    Application.OnTime startTime, "getCAP"

End Sub


End Sub
