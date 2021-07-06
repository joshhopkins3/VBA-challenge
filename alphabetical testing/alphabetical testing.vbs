Attribute VB_Name = "Module1"
Sub marketanalysis()

Dim ticker As String
Dim yearlychange
Dim percentchange
Dim totalstock As String

 'Inserting Data Via Ranges
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"

 Dim last as long
last = Cells(Rows.Count, 1).End(xlUp).Row

Dim i

For i = 2 To last

ticker = Cells(i, 1).Value

Cells(i, 9).Value = ticker

Next

For i = 2 To last

yearlychange = (Cells(i, 6).Value - Cells(i, 3).Value)

Cells(i, 10).Value = yearlychange

Next


For i = 2 To last

percentchange = (Cells(i, 6).Value / Cells(i, 3).Value) - 1

Cells(i, 11).NumberFormat = "0.00%"

Cells(i, 11).Value = percentchange


Next

For i = 2 To last

totalstock = Cells(i, 7).Value

Cells(i, 12).Value = totalstock



Next


End Sub
