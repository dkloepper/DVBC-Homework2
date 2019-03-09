Attribute VB_Name = "Module1"
Option Explicit


Sub summarizeStocks()

    'Initalize variables
    Dim ws, wsAdd As Worksheet
    Dim i As Long
    Dim s, t As Integer
    
    Dim endRow As Long
    
    Dim ticker As String
    
    Dim volume As Double
    Dim openPrice, closePrice As Double
    
    Dim yearChange, yearPctChange As Double
    
    Dim grtPctIncr, grtPctDecr, grtVolume As Double
    Dim grtPctIncrTicker, grtPctDecrTicker, grtVolumeTicker
    
    'Loop through each worksheet in the workbook
    For Each ws In ThisWorkbook.Sheets
        
        'Reset default starting values for the loop
        s = 2
        openPrice = Null
    
        'Get year from worksheet name
        Dim year As String
        year = ws.Name
        
        'Create a new summary worksheet for each year
        Set wsAdd = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsAdd.Name = year + " Summary"
        
        'Add column headings to new worksheet
        wsAdd.Range("A1").Value = "Ticker"
        wsAdd.Range("B1").Value = "Yearly Change"
        wsAdd.Range("C1").Value = "Percent Change"
        wsAdd.Range("D1").Value = "Volume"
        wsAdd.Range("A1:D1").Font.Bold = True
        
        'Bring focus back to the worksheet with data
        ws.Activate
        
        'Get last row containg data for the next loop
        endRow = Cells(Rows.Count, 1).End(xlUp).Row
        
        'Sort the data by the ticker and timestamp
        Range("A1:G" & endRow).Sort key1:=Range("A1:A" & endRow), order1:=xlAscending, key2:=Range("B1:B" & endRow), order2:=xlAscending, Header:=xlYes
        
        'Loop through data rows, starting with row 2 to maximum row value
        For i = 2 To endRow
        
            ticker = Cells(i, 1).Value
            'Collect the opening price, but only if there isn't already a value set
            If IsNull(openPrice) Then
                openPrice = Cells(i, 3).Value
            Else
                'Do Nothing
            End If
            'Check if ticker value in next row is the same
            If Cells(i + 1, 1).Value <> ticker Then
                'If ticker value is different in the next row down:
                
                'Add the last volume value to the total
                volume = volume + Cells(i, 7).Value
                
                'Collect the closing price
                closePrice = Cells(i, 6).Value
                
                'Calculate the yearly change
                yearChange = closePrice - openPrice
                
                'Calculate the yearly percentage change, setting defaults if the denominator will be zero
                If openPrice = 0 Then
                    If closePrice = 0 Then
                        yearPctChange = 0
                    Else
                        yearPctChange = 1
                    End If
                Else
                    yearPctChange = yearChange / openPrice
                End If
                
                'Populate the summary values into the appropriate cells in the summary worksheet
                wsAdd.Cells(s, 1).Value = ticker
                wsAdd.Cells(s, 2).Value = Format(yearChange, "#,##0.0#")
                wsAdd.Cells(s, 3).Value = Format(yearPctChange, "##0.00%")
                wsAdd.Cells(s, 4).Value = Format(volume, "#,##0")
                
                'Iterate the value for the next summary row
                s = s + 1
                
                'Reset values for the next loop
                volume = 0
                openPrice = Null
                closePrice = 0
            Else
                'If ticker is the same, continue to add to the volume total
                volume = volume + Cells(i, 7).Value
            End If
        
        Next i
        
        'Add conditional formating to the yearly change column
        With wsAdd.Range("B:B").FormatConditions.Add(xlCellValue, xlGreater, 0)
            .Interior.Color = vbGreen
        End With
        
        With wsAdd.Range("B:B").FormatConditions.Add(xlCellValue, xlLess, 0)
            .Interior.Color = vbRed
        End With
    
        'Add column and row headings for summarization
        wsAdd.Range("G2").Value = "Greatest % Increase"
        wsAdd.Range("G3").Value = "Greatest % Decrease"
        wsAdd.Range("G4").Value = "Greatest Total Volume"
        wsAdd.Range("H1").Value = "Ticker"
        wsAdd.Range("I1").Value = "Value"
        wsAdd.Range("G1:G4").Font.Bold = True
        wsAdd.Range("G1:I1").Font.Bold = True
        
        'Reset any previous values
        grtPctIncr = 0
        grtPctDecr = 0
        grtVolume = 0
        grtPctIncrTicker = ""
        grtPctDecrTicker = ""
        grtVolumeTicker = ""
        
        'Loop through the summary data to determine the tickers with greatest pct increase, pct decrease and volume
        For t = 2 To wsAdd.Cells(Rows.Count, 1).End(xlUp).Row
            If wsAdd.Cells(t, 3).Value > grtPctIncr Then
                grtPctIncr = wsAdd.Cells(t, 3).Value
                grtPctIncrTicker = wsAdd.Cells(t, 1).Value
            End If
            If wsAdd.Cells(t, 3).Value < grtPctDecr Then
                grtPctDecr = wsAdd.Cells(t, 3).Value
                grtPctDecrTicker = wsAdd.Cells(t, 1).Value
            End If
            If wsAdd.Cells(t, 4).Value > grtVolume Then
                grtVolume = wsAdd.Cells(t, 4).Value
                grtVolumeTicker = wsAdd.Cells(t, 1).Value
            End If
        Next t
        
        'Insert ticker symbols
        wsAdd.Range("H2").Value = grtPctIncrTicker
        wsAdd.Range("H3").Value = grtPctDecrTicker
        wsAdd.Range("H4").Value = grtVolumeTicker
        
        'Insert Incr, Decr and Volume for each ticker
        wsAdd.Range("I2").Value = Format(grtPctIncr, "#.00%")
        wsAdd.Range("I3").Value = Format(grtPctDecr, "#.00%")
        wsAdd.Range("I4").Value = Format(grtVolume, "#,##0")
        
        'Autofit the column sizes on the summary worksheet
        wsAdd.Columns("A:I").AutoFit
        
    Next


End Sub
