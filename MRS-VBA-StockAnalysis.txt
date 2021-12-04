
Sub stockAnalysis()
    'VARIABLE FOR WORKSHEET
    Dim ws          As Worksheet
    'VARIABLE FOR SUMMARY ROW
    Dim sumRow      As Integer
    'VARIABLE FOR STOCK SYMBOL
    Dim tickerSymbol As String
    'VARIABLE FOR STOCK OPEN VALUE
    Dim openVal     As Double
    'VARIABLE FOR STOCK CLOSE VALUE
    Dim closeVal    As Double
    'VARIABLE FOR YEARLY CHANGE (VALUE)
    Dim yearlyChange As Double
    'VARIABLE FOR VOLUME
    Dim volume      As LongLong
    'VARIABLE FOR YEARLY PERCENT CHANGE
    Dim pctChange   As Double
    
    'FOR EACH LOOP OF WORKSHEETS
    For Each ws In ActiveWorkbook.Worksheets
        ws.Activate
        
        'START VALUES
        iValue = 1
        sumRow = 2
        volume = 0
        
        Range("i1:l1").Interior.ColorIndex = 1
        Range("i1:l1").Font.ColorIndex = 2
        Range("i1:l1").Font.Bold = True
        Range("i1").Value = "Symbol"
        Range("j1").Value = "Year Change"
        Range("k1").Value = "% Change"
        Range("l1").Value = "Total Volume"
        
        'DEFINE END ROW
        endRow = Cells(Rows.Count, "A").End(xlUp).Row
        
        For i = 2 To endRow
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                
                tickerSymbol = Cells(i, 1).Value
                iValue = iValue + 1
                
                openVal = Cells(iValue, 3).Value
                closeVal = Cells(i, 6).Value
                
                For j = iValue To i
                    volume = volume + Cells(j, 7).Value
                    
                Next j
                
                If openVal = 0 Then
                    pctChange = closeVal
                    
                Else
                    yearlyChange = closeVal - openVal        '
                    pctChange = yearlyChange / openVal
                    
                End If
                
                'CREATE SUMMARY TABLE
                Cells(sumRow, 9).Value = tickerSymbol
                Cells(sumRow, 10).Value = yearlyChange
                Cells(sumRow, 11).Value = pctChange
                Cells(sumRow, 12).Value = volume
                
                sumRow = sumRow + 1
                volume = 0
                yearlyChange = 0
                pctChange = 0
                iValue = i
            End If
        Next i
        'FORMAT BASED ON POSITIVE OR NEGATIVE YEARLY CHANGE
        newEndRow = Cells(Rows.Count, "J").End(xlUp).Row
        For j = 2 To newEndRow
            If Cells(j, 10) > 0 Then
                Cells(j, 10).Interior.ColorIndex = 50
                Cells(j, 10).Font.ColorIndex = 2
            Else
                Cells(j, 10).Interior.ColorIndex = 9
                Cells(j, 10).Font.ColorIndex = 2
            End If
        Next j
        
        Columns("i:l").AutoFit
        Columns("i:l").Font.Bold = True
        With Range("i:l")
            With .Borders
                .LineStyle = xlContinuous
                .Color = vbBlack
                .Weight = xlThin
            End With
        End With
        Columns("j:j").Style = "Currency"
        Columns("k:k").NumberFormat = "0.00%"
        Columns("l:l").NumberFormat = "#,##0"
        
    Next ws
    
End Sub
