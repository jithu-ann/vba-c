VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub ticker()


Dim tickername As String
Dim tickervolume As Double
tickervolume = 0
Dim summaryticker As Integer
summaryticker = 2
Dim openingprice As Double
openingprice = Cells(2, 3).Value
Dim closingprice As Double
Dim YearlyChange As Double
Dim percentchange As Double
Dim GreatestIncrease As Double
Dim GreatestDecrease As Double
Dim GreatestTotal As Double
Dim GreatestIncreaseTicker As String
Dim GreatestDecreaseTicker As String
Dim GreatestTotalTicker As String
'table headings
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"
Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"

Dim LastRow As Long
LastRow = Cells(Rows.Count, "A").End(xlUp).Row
For i = 2 To LastRow
If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
'ticker names
tickername = Cells(i, 1).Value
tickervolume = tickervolume + Cells(i, 7).Value
Range("I" & summaryticker).Value = tickername
Range("L" & summaryticker).Value = tickervolume
closingprice = Cells(i, 6).Value
YearlyChange = (closingprice - openingprice)
Range("J" & summaryticker).Value = YearlyChange
If (openingprice = 0) Then
    percentchange = 0
Else
    percentchange = YearlyChange / openingprice
End If
Range("K" & summaryticker).Value = percentchange
Range("k" & summaryticker).NumberFormat = "0.00%"
summaryticker = summaryticker + 1
tickervolume = 0
openingprice = Cells(i + 1, 3)
Else
    tickervolume = tickervolume + Cells(i, 7).Value
End If
Next i
lastrow_summarytable = Cells(Rows.Count, 9).End(xlUp).Row
    
    'Color code yearly change
    
    For i = 2 To lastrow_summarytable
            If Cells(i, 10).Value > 0 Then
                Cells(i, 10).Interior.ColorIndex = 10
            Else
                Cells(i, 10).Interior.ColorIndex = 3
            End If
    Next i
    LastRow = Cells(Rows.Count, "I").End(x1up).Row
    GreatestIncrease = 0
    GreatestDecrease = 0
    GreatestTotal = 0
    GreatestIncreaseTicker = ""
    GreatestDecreaaseTicker = ""
    GreatestTotalTicker = ""
    For i = 2 To LastRow
    
     If Cells(i, 11).Value > GreatestIncrease Then
            GreatestIncreaseTicker = Cells(i, 9).Value
            GreatestIncrease = Cells(i, 11).Value
        End If
        
        If Cells(i, 11).Value < GreatestDecrease Then
            GreatestDecreaseTicker = Cells(i, 9).Value
            GreatestDecrease = Cells(i, 11).Value
        End If
        
        If Cells(i, 12).Value > GreatestTotal Then
            GreatestTotalTicker = Cells(i, 9).Value
            GreatestTotal = Cells(i, 12).Value
        End If
         
    Next
    
    
    Cells(2, 16).Value = GreatestIncreaseTicker
    Cells(2, 17).Value = GreatestIncrease
    Cells(2, 17).NumberFormat = "0.00%"
    Cells(3, 16).Value = GreatestDecreaseTicker
    Cells(3, 17).Value = GreatestDecrease
    Cells(3, 17).NumberFormat = "0.00%"
    Cells(4, 16).Value = GreatestTotalTicker
    Cells(4, 17).Value = GreatestTotal
    Cells(4, 17).NumberFormat = "#"

   
    

End Sub
    

