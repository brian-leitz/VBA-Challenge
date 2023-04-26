# VBA-Challenge
Working with stock data using VBA
Google Drive Link for Module 2 Challenge Submission

https://drive.google.com/drive/folders/1sss_QY2zjM5L7JTGPkgeaQ0mWkkT9rEX?usp=share_link


Code Written: 

```

Sub Stock_Analysis():

Dim lastrow As Long
Dim Ticker As String
Dim openprice As Double
Dim closeprice As Double
Dim yearlychange As Double
Dim percentagechange As Double
Dim totalvolume As LongLong
Dim yearlyrow As Integer

For Each ws In Worksheets

ws.Activate


yearlyrow = 2
lastrow = Cells(Rows.Count, 1).End(xlUp).Row
openprice = Cells(2, 3).Value
totalvolume = 0

    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    
    
For Myrow = 2 To lastrow

totalvolume = totalvolume + Cells(Myrow, 7)
    
If Cells(Myrow + 1, 1).Value <> Cells(Myrow, 1).Value Then
    
    Ticker = Cells(Myrow, 1).Value
    closeprice = Cells(Myrow, 6).Value
    yearlychange = closeprice - openprice
    percentagechange = yearlychange / openprice
    
    Cells(yearlyrow, 9).Value = Ticker
    Cells(yearlyrow, 10).Value = yearlychange
    Cells(yearlyrow, 11).Value = percentagechange
    Cells(yearlyrow, 12).Value = totalvolume
    
    If yearlychange >= 0 Then
    
        Cells(yearlyrow, 10).Interior.ColorIndex = 4
        
        Else
        
        Cells(yearlyrow, 10).Interior.ColorIndex = 3
     End If
        
    
    yearlyrow = yearlyrow + 1
   
    openprice = Cells(Myrow + 1, 3).Value
    totalvolume = 0
    
End If

Next Myrow
    Range("K2:K" & yearlyrow).Style = "Percent"

Next ws

End Sub
```
