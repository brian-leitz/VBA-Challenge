Attribute VB_Name = "Module1"
Sub Stock_Analysis_Test():

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

lastrow = Cells(Rows.Count, 1).End(xlUp).Row
yearlyrow = 2
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
    percentagechange = closeprice / openprice
    
    Cells(yearlyrow, 9).Value = Ticker
    Cells(yearlyrow, 10).Value = yearlychange
    Cells(yearlyrow, 11).Value = percentagechange
    Cells(yearlyrow, 12).Value = totalvolume
    
    If yearlychange >= 0 Then
    
        Cells(yearlyrow, 10).Interior.ColorIndex = 4
        
        Else
        
        Cells(yearlyrow, 10).Interior.ColorIndex = 3
        
        
    
    yearlyrow = yearlyrow + 1
   
    openprice = Cells(Myrow + 1, 3).Value
    totalvolume = 0
        
    End If
    
End If

Next Myrow

    Range("K2:K" & yearlyrow).Style = "Percent"

Next ws

End Sub
