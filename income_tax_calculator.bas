Attribute VB_Name = "Module4"
Option Explicit
Function tax(x As Double) As Variant
Dim tempArr As Variant
Dim i As Integer
ReDim tempArr(1 To 7, 1 To 3)
tempArr(1, 1) = "General Income Slab"
tempArr(1, 2) = "Taxable Income"
tempArr(1, 3) = "Tax Amount"
tempArr(2, 1) = "Upto 4 lakhs"
tempArr(3, 1) = "4 to 5 lakhs"
tempArr(4, 1) = "5 to 7 lakhs"
tempArr(5, 1) = "7 to 20 lakhs"
tempArr(6, 1) = "Over 20 lakhs"
tempArr(7, 1) = "Sum"
tempArr(7, 2) = x
tempArr(7, 3) = 0

If x >= 2000000 Then
    tempArr(2, 2) = 400000
    tempArr(2, 3) = 4000
    tempArr(3, 2) = 100000
    tempArr(3, 3) = 10000
    tempArr(4, 2) = 200000
    tempArr(4, 3) = 40000
    tempArr(5, 2) = 1300000
    tempArr(5, 3) = 390000
    tempArr(6, 2) = x - 2000000
    tempArr(6, 3) = (x - 2000000) * 0.36
    For i = 2 To 6
        tempArr(7, 3) = tempArr(i, 3) + tempArr(7, 3)
    Next i
    
    'tax = (x - 2000000) * 0.36 + 444000
ElseIf x >= 700000 Then
    tempArr(2, 2) = 400000
    tempArr(2, 3) = 4000
    tempArr(3, 2) = 100000
    tempArr(3, 3) = 10000
    tempArr(4, 2) = 200000
    tempArr(4, 3) = 40000
    tempArr(5, 2) = (x - 700000)
    tempArr(5, 3) = (x - 700000) * 0.3
    tempArr(6, 2) = "-"
    tempArr(6, 2) = "-"
    For i = 2 To 5
        tempArr(7, 3) = tempArr(i, 3) + tempArr(7, 3)
    Next i
    'tax = (x - 700000) * 0.3 + 54000
ElseIf x >= 500000 Then
    tempArr(2, 2) = 400000
    tempArr(2, 3) = 4000
    tempArr(3, 2) = 100000
    tempArr(3, 3) = 10000
    tempArr(4, 2) = (x - 500000)
    tempArr(4, 3) = (x - 500000) * 0.2
    tempArr(5, 2) = "-"
    tempArr(5, 3) = "-"
    tempArr(6, 2) = "-"
    tempArr(6, 3) = "-"
    For i = 2 To 4
        tempArr(7, 3) = tempArr(i, 3) + tempArr(7, 3)
    Next i
    'tax = (x - 500000) * 0.2 + 14000
ElseIf x >= 400000 Then
    tempArr(2, 2) = 400000
    tempArr(2, 3) = 4000
    tempArr(3, 2) = (x - 400000)
    tempArr(3, 3) = (x - 400000) * 0.1
    tempArr(4, 2) = "-"
    tempArr(4, 3) = "-"
    tempArr(5, 2) = "-"
    tempArr(5, 3) = "-"
    tempArr(6, 2) = "-"
    tempArr(6, 3) = "-"
    For i = 2 To 3
        tempArr(7, 3) = tempArr(i, 3) + tempArr(7, 3)
    Next i
    'tax = (x - 400000) * 0.1 + 4000
Else
    tempArr(2, 2) = x
    tempArr(2, 3) = x * 0.01
    tempArr(3, 2) = "-"
    tempArr(3, 3) = "-"
    tempArr(4, 2) = "-"
    tempArr(4, 3) = "-"
    tempArr(5, 2) = "-"
    tempArr(5, 3) = "-"
    tempArr(6, 2) = "-"
    tempArr(6, 3) = "-"
    tempArr(7, 3) = x * 0.01
End If
tax = tempArr
End Function
