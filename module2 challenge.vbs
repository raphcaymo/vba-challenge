Attribute VB_Name = "Module1"
Sub getSummary()
  Dim ws As Worksheet
 
  For Each ws In Worksheets
    
    'defining variables. I used LongLong data type to volCounter and summMaxVolume due to carry big numbers of integers
        Dim volCounter, summMaxVol As LongLong
        Dim i, j, lastRow, summRow As Long
        Dim OpPrice, ClPrice, PriceChange, PriceChangePer, summIncPer, summDecPer As Double
        Dim TickerName, summIncTicker, summDecTicker, summVolTicker As String
            lastRow = Cells(Rows.Count, 1).End(xlUp).Row


   'Will setup header needed
            ws.Range("I1").Value = "Ticker"
            ws.Range("J1").Value = "Yearly Change"
            ws.Range("K1").Value = "Percentage Change"
            ws.Range("L1").Value = "Total Stock Volume"
            
    'Setting up default value for the first Opening price of the first ticker name and as well as for variable summRow
            OpPrice = ws.Cells(2, 3).Value
            summRow = 2
    
    'for loop starts from here. This for loop says that as long as two are with the same ticker name, the volume will keep adding. If not, it will catch the ticker name, opening, closing, and calculate price change, and price change rate
            For i = 2 To lastRow
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                    TickerName = ws.Cells(i, 1).Value
                    ClPrice = ws.Cells(i, 6).Value
                    PriceChange = ClPrice - OpPrice
                    
    'for as long as OpPrice is not 0 it will calculate the Price change rate
                    If OpPrice <> 0 Then
                        PriceChangePer = PriceChange / OpPrice
                    End If
                    
                    volCounter = volCounter + ws.Cells(i, 7).Value
                    
   'this line will write values in Columns I to L, it involves cell formatting and styling as well
                    ws.Range("I" & summRow).Value = TickerName
                    ws.Range("J" & summRow).Value = PriceChange
                    ws.Range("J" & summRow).NumberFormat = "$0.00"
                    
                    If (PriceChange > 0) Then
                        ws.Range("J" & summRow).Interior.ColorIndex = 4
                    ElseIf (PriceRange <= 0) Then
                        ws.Range("J" & summRow).Interior.ColorIndex = 3
                    End If
                    
                    ws.Range("K" & summRow).Value = PriceChangePer
                    ws.Range("K" & summRow).NumberFormat = "0.00%"
                    ws.Range("L" & summRow).Value = volCounter
                                   
    'var summRow will be ready'd for the next sequence of I. In the first part of the program it has default value of 2. Similar with var ClPrice and PriceChange
                    summRow = summRow + 1
                    PriceChange = 0
                    ClPrice = 0
                    OpPrice = ws.Cells(i + 1, 3).Value
              
    'Catching Greatest Increase, Decrease, and Volume Values
                    If (PriceChangePer > summIncPer) Then
                        summIncPer = PriceChangePer
                        summIncTicker = TickerName
                    ElseIf (PriceChangePer < summDecPer) Then
                        summDecPer = PriceChangePer
                        summDecTicker = TickerName
                    End If
    
                    If (volCounter > summMaxVol) Then
                        summMaxVol = volCounter
                        summVolName = TickerName
                    End If
    
                    PriceChangePer = 0
                    volCounter = 0
                Else
                    volCounter = volCounter + ws.Cells(i, 7).Value
             End If
             
        Next i
        
        'Greatest % Inc and % Dec and Total Volume
        ws.Range("O1").Value = "Ticker"
        ws.Range("P1").Value = "Value"
        
        ws.Range("N2").Value = "Greatest % Increase"
        ws.Range("O2").Value = summIncTicker
        ws.Range("P2").Value = summIncPer
        
        ws.Range("N3").Value = "Greatest % Decrease"
        ws.Range("O3").Value = summDecTicker
        ws.Range("P3").Value = summDecPer
        
        ws.Range("P2:P3").NumberFormat = "0.00%"
        
        ws.Range("N4").Value = "Greatest Total Volume"
        ws.Range("O4").Value = summVolName
        ws.Range("P4").Value = summMaxVol
        

    'this line will auto adjust each columns to better fit data in each cells
        ws.Columns("A:Q").AutoFit
        
    'this will run across worksheets
    Next ws
    
    'this code is indication that the program is already done.
    MsgBox ("Workbook summarized.")


End Sub
