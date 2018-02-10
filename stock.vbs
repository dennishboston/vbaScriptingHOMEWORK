Sub stock()
Dim firstRow As Integer
Dim Ticker As String
Dim TotalVolume As Long

firstRow = 2

Range("I1").Value = "Ticker"
Range("j1").Value = "TotalVolume"


For i = 2 To 1000000
If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

Ticker = Cells(i, 1).Value
Range("I" & firstRow) = Ticker
TotalVolume = TotalVolume + Cells(i, 7).Value
Range("J" & firstRow) = TotalVolume
TotalVolume = 0
firstRow = firstRow + 1

Else
TotalStock = TotalStock + Cells(i, 7).Value
               
End If

Next i

End Sub


