Sub Ticker()

Dim cWs As Worksheet 'Defining current worksheet
Dim cWb As Workbook   'Define current workbook
Dim newWs As Worksheet
Dim wsNames() As Variant ' array containing current Worksheet names
Dim currentSheetCount As Integer ' Current Sheet Count
Dim wsName As String ' Worksheet name
Dim OpeningPrice As Double
Dim tickerRangeStart As Long
Dim tickerRangeEnd As Long
Dim YearlyChange As Double
Dim TickerCounter As Long

Dim ValueRange As Range
Dim NameRange As Range
Dim NameCheck As String

TickerCounter = 0

Set cWb = Application.ActiveWorkbook
currentSheetCount = cWb.Sheets.Count
ReDim wsNames(cWb.Sheets.Count - 1)

OriginalSheetCount = cWb.Sheets.Count

For i = 1 To cWb.Sheets.Count
    wsName = cWb.Sheets(i).Name
    wsNames(i - 1) = wsName
Next

For i = 1 To OriginalSheetCount
    Set newWs = cWb.Worksheets.Add(Before:=Worksheets(1))
    newWs.Name = "My Summary" & "_" & wsNames(i - 1)
Next i

''' Main LOOP


For Each ws In cWb.Worksheets

''' Skip summary sheets

NameCheck = Left(ws.Name, 10)

    If NameCheck <> "My Summary" Then
    
        tickerRangeStart = 0
        StockVolume = 0
        TickerCounter = 1
        
        Set cWs = ws
        cWs.Activate ' Activate current worksheet
        
        cSummaryName = "My Summary" & "_" & cWs.Name
        tickerRangeStart = 2
        
        cWb.Worksheets(cSummaryName).Columns("K").ColumnWidth = 24
        
        cWb.Worksheets(cSummaryName).Cells(2, 11).Value = "Best Performing Stock"
        cWb.Worksheets(cSummaryName).Cells(2, 11).Interior.ColorIndex = 43
        
        
        cWb.Worksheets(cSummaryName).Cells(3, 11).Value = "Worst Performing Stock"
        cWb.Worksheets(cSummaryName).Cells(3, 11).Interior.ColorIndex = 45
        
        cWb.Worksheets(cSummaryName).Cells(4, 11).Value = "Greatest Total Value"
        cWb.Worksheets(cSummaryName).Cells(4, 11).Interior.ColorIndex = 47
        
         cWb.Worksheets(cSummaryName).Cells(1, 1).Value = "<Ticker>"
         cWb.Worksheets(cSummaryName).Cells(1, 2).Value = "<Year Open>"
         cWb.Worksheets(cSummaryName).Cells(1, 3).Value = "<Year Close>"
         cWb.Worksheets(cSummaryName).Cells(1, 4).Value = "<Year Change>"
         cWb.Worksheets(cSummaryName).Cells(1, 6).Value = "<Year Growth>"
         cWb.Worksheets(cSummaryName).Cells(1, 7).Value = "<Year Volume>"
        
        LastRow = cWs.Cells(Rows.Count, 1).End(xlUp).Row
        For j = 2 To LastRow
            
            YearlyChange = 0 ' Reset / Set yearly change
            MaxOpeningPrice = 0 ' Reset / Set Max opening price
            MinOpeningPrice = 0 ' Reset / Set Min opening price
            
            
            tickerRangeEnd = j
            StockVolume = StockVolume + cWs.Cells(j, 7).Value
            
            If Cells(j + 1, 1).Value <> Cells(j, 1).Value Then
            
                TickerCounter = TickerCounter + 1
            
                cWb.Worksheets(cSummaryName).Cells(TickerCounter, 1).Value = Cells(j, 1).Value
                
                
                MaxPrice = Application.WorksheetFunction.Max(cWs.Range("C" & tickerRangeStart & ":C" & tickerRangeEnd))
                MinPrice = Application.WorksheetFunction.Min(cWs.Range("C" & tickerRangeStart & ":C" & tickerRangeEnd))
                
                BOYopen = cWs.Cells(tickerRangeStart, 3).Value ' get value of stock at opening - begining of the year
                EOYclose = cWs.Cells(tickerRangeEnd, 6).Value  ' get value of stock by closing - end of the year
                                
                YearlyChange = EOYclose - BOYopen
                YearlyGrowth = Growth(BOYopen, YearlyChange)
                
                ' Start writing values into Summary Sheets
                cWb.Worksheets(cSummaryName).Cells(TickerCounter, 2).Value = BOYopen
                cWb.Worksheets(cSummaryName).Cells(TickerCounter, 3).Value = EOYclose
                cWb.Worksheets(cSummaryName).Cells(TickerCounter, 4).Value = YearlyChange
                cWb.Worksheets(cSummaryName).Cells(TickerCounter, 6).Value = YearlyGrowth
                cWb.Worksheets(cSummaryName).Cells(TickerCounter, 7).Value = StockVolume
                             
                
                
                StockVolume = 0 ' reset stock volume
                tickerRangeStart = j + 1
            End If
                
                
            If TickerCounter > 0 And YearlyChange > 0 Then
                cWb.Worksheets(cSummaryName).Cells(TickerCounter, 6).Interior.ColorIndex = 43
               ElseIf TickerCounter > 0 And YearlyChange < 0 Then
                cWb.Worksheets(cSummaryName).Cells(TickerCounter, 6).Interior.ColorIndex = 45
            End If
        
        Next j
    


    
    Set cWs = cWb.Worksheets(cSummaryName)
    cWs.Activate
    'cWb.Worksheets("My Summary").Range("F1:F" & TickerCounter).NumberFormat = "0.00%"
    cWs.Range("F1:F" & TickerCounter).NumberFormat = "0.00%"
    cWs.Range("L2:L3").NumberFormat = "0.00%"
    LastRow = cWs.Cells(Rows.Count, 1).End(xlUp).Row
    
    Set ValueRange = cWs.Range("F2:F" & LastRow)
    Set VolumeRange = cWs.Range("G2:G" & LastRow)
    Set NameRange = cWs.Range("A2:A" & LastRow)
    
    Best_Stock = Application.WorksheetFunction.Max(ValueRange)
    Result = Application.WorksheetFunction.Index(NameRange, Application.WorksheetFunction.Match(Best_Stock, ValueRange, 0))
    cWs.Cells(2, 12).Value = Best_Stock
    cWs.Cells(2, 13).Value = Result
    
    
    Worst_Stock = Application.WorksheetFunction.Min(ValueRange)
    Result = Application.WorksheetFunction.Index(NameRange, Application.WorksheetFunction.Match(Worst_Stock, ValueRange, 0))
    cWs.Cells(3, 12).Value = Worst_Stock
    cWs.Cells(3, 13).Value = Result
    
    MaxVolume_Stock = Application.WorksheetFunction.Max(VolumeRange)
    Result = Application.WorksheetFunction.Index(NameRange, Application.WorksheetFunction.Match(MaxVolume_Stock, VolumeRange, 0))
    cWs.Cells(4, 12).Value = MaxVolume_Stock
    cWs.Cells(4, 13).Value = Result
    
    cWs.Range("A:M").EntireColumn.AutoFit
    
    End If
Next
End Sub

Private Function Growth(inBOY As Variant, inChange As Variant)
        If inBOY <> 0 Then
            Growth = inChange / inBOY
            Else
            Growth = "N / A"
        End If
End Function

Private Function AddressOfMax(rng As Range) As Range ' Stack Overflow Code
    Set AddressOfMax = rng.Cells(WorksheetFunction.Match(WorksheetFunction.Max(rng), rng, 0))
End Function
Private Function AddressOfMin(rng As Range) As Range ' Stack Overflow Code
    Set AddressOfMax = rng.Cells(WorksheetFunction.Match(WorksheetFunction.Mim(rng), rng, 0))
End Function

''LEFTOVER Code
'MaxOpeningPrice = Application.WorksheetFunction.Max(cWs.Range(cWs.Cells(tickerRangeStart, 3), cWs.Cells(tickerRangeEnd, 3)))
'MinOpeningPrice = Application.WorksheetFunction.Min(cWs.Range(cWs.Cells(tickerRangeStart, 3), cWs.Cells(tickerRangeEnd, 3)))
'' NOT USED
Private Function ValueExists(testedValue As Variant, arr As Variant) As Boolean
Dim memeber As Variant
On Error GoTo IsInArrayError: 'array is empty
    For Each member In arr
        If member = testedValue Then
            ValueExists = True
            Exit Function
        End If
    Next member
Exit Function
IsInArrayError:
On Error GoTo 0
ValueExists = False
End Function



