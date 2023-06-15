Attribute VB_Name = "Module1"
Sub challengefinal()
    Dim ws As Worksheet
    Dim wb As Workbook
    Dim lastrow As Long
    Dim columnA As Range
    Dim ticker As Range
    Dim outputRow As Long
    Dim tickerDict As Object
    Dim maxpercent As Double
    Dim maxpercentticker As String
    Dim minpercent As Double
    Dim minpercentticker As String
    Dim maxvolume As Double
    Dim maxvolumeticker As String
    
    Set wb = ThisWorkbook
    For Each ws In wb.Worksheets
            'Find the last row of data in the worksheet
            lastrow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            
            ' Set the range for the ticker symbols column
            Set columnA = ws.Range("A2:A" & lastrow)
            
            ' Set up dictionary with unique values
            Set tickerDict = CreateObject("Scripting.Dictionary")
            For Each ticker In columnA
                 tickerDict(ticker.Value) = 1
            Next ticker
            
            ' Start the output row with room for headers
            outputRow = 2
            
            ' Start variables for tracking max/min values
            maxpercent = -99999999
            minpercent = 99999999
            maxvolume = 0
            
            Dim tickersymbol As Variant
            For Each tickersymbol In tickerDict.keys
            
                'find first and last row for ticker symbol
                Dim firstinstance As Long
                Dim lastinstance As Long
                firstinstance = ws.Range("A:A").Find(tickersymbol).Row
                lastinstance = ws.Range("A:A").Find(tickersymbol, ws.Cells(firstinstance, 1), xlValues, xlWhole, xlByRows, xlPrevious).Row
                
                'calculate symbol summary values
                Dim openingPrice As Double
                Dim closingPrice As Double
                Dim yearlyChange As Double
                Dim percentChange As Double
                Dim totalVolume As Double
    
                openingPrice = ws.Cells(firstinstance, 3).Value
                closingPrice = ws.Cells(lastinstance, 6).Value
    
                yearlyChange = closingPrice - openingPrice
                percentChange = yearlyChange / openingPrice
                totalVolume = WorksheetFunction.Sum(ws.Range("G" & firstinstance & ":G" & lastinstance))
                
                'Evaluate and update summary values
                If percentChange > maxpercent Then
                    maxpercent = percentChange
                    maxpercentticker = tickersymbol
                End If
    
                If percentChange < minpercent Then
                     minpercent = percentChange
                     minpercentticker = tickersymbol
                End If
    
                If totalVolume > maxvolume Then
                    maxvolume = totalVolume
                    maxvolumeticker = tickersymbol
                End If
                
                ' Output the results in the adjacent columns (And format %change column)
                ws.Cells(outputRow, 9).Value = tickersymbol
                ws.Cells(outputRow, 10).Value = yearlyChange
                ws.Cells(outputRow, 11).Value = percentChange
                ws.Cells(outputRow, 11).NumberFormat = "0.00%"
                ws.Cells(outputRow, 12).Value = totalVolume
    
                'Conditional Formatting
                If ws.Cells(outputRow, 10).Value >= 0 Then
                    ws.Cells(outputRow, 10).Interior.ColorIndex = 4
                    ws.Cells(outputRow, 11).Interior.ColorIndex = 4
                Else
                    ws.Cells(outputRow, 10).Interior.ColorIndex = 3
                    ws.Cells(outputRow, 11).Interior.ColorIndex = 3
                End If
                
                ' Move to the next output row
                outputRow = outputRow + 1
            Next tickersymbol
            
            'output summary values and format columns
            ws.Range("N2").Value = "Max % gain"
            ws.Range("N3").Value = "Max % loss"
            ws.Range("N4").Value = "Max volume"
            ws.Range("P2").Value = maxpercent
            ws.Range("P3").Value = minpercent
            ws.Range("P2").NumberFormat = "0.00%"
            ws.Range("P3").NumberFormat = "0.00%"
            ws.Range("P4").Value = maxvolume
            ws.Range("O2").Value = maxpercentticker
            ws.Range("O3").Value = minpercentticker
            ws.Range("O4").Value = maxvolumeticker
            ws.Range("I1").Value = "Ticker"
            ws.Range("J1").Value = "Yearly Change"
            ws.Range("K1").Value = "Percent Change"
            ws.Range("L1").Value = "Total Stock Volume"
            ws.UsedRange.EntireColumn.AutoFit
            ws.UsedRange.EntireRow.AutoFit
    Next ws
       
End Sub



