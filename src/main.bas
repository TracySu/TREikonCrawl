' ===============================================
' Author: Yuanjia Su <yuanjia.su.13 at ucl.ac.uk>
'
' Description:
'         1. loop over rows on 'sheet2'
' ===============================================


Sub main()
    With Worksheets("Sheet2")
        
        Debug.Print "Start: " & Now
        
        Call formulas
        Call calcmetrics(2, 2, 1)
    
    End With
End Sub

