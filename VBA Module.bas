Attribute VB_Name = "Module1"
Sub ProcessAll()
Dim x, y As Integer
For x = 1 To Worksheets.Count
    Worksheets(x).Select
    Call Title
    Call Ticker
    Call Change
    Call Color
    Call MinMax
    Call GreatVolume
    Call Format
Next x

End Sub
Sub Title()
[I1].Value = "Ticker"
[J1].Value = "Yearly Change"
[K1].Value = "Percentage Change"
[L1].Value = "Total Stock Volume"
[I1:L1].Font.Bold = True
Columns("I:L").AutoFit
End Sub
Sub Ticker()
Dim x, y As Integer
x = 1
y = 2
Do While Cells(x, 1).Value <> ""
    x = x + 1
    If Cells(x, 1).Value <> Cells(x - 1, 1).Value Then
    Cells(y, 9).Value = Cells(x, 1).Value
    y = y + 1
    Else
    End If
         
Loop

End Sub

Sub Change()
Dim x, y, a, b As Integer
Dim op, cl, percent As Double
x = 1
y = 2
cl = 0
op = 0
Do Until IsEmpty(Cells(x, 1).Value)
    x = x + 1
    If Cells(x, 1).Value <> Cells(x - 1, 1).Value Then
    op = Cells(x, 3).Value
    a = x
    ElseIf Cells(x, 1).Value <> Cells(x + 1, 1).Value Then
    cl = Cells(x, 6).Value
    Cells(y, 10).Value = cl - op
        If op > 0 Then
        Cells(y, 11).Value = (cl - op) / op
        Cells(y, 11).NumberFormat = "0.00%"
        Else
        Cells(y, 11).Value = ""
        End If
    Cells(y, 12).Value = Application.Sum(Range(Cells(a, 7), Cells(x, 7)))
' sum: https://stackoverflow.com/questions/42844778/vba-for-each-cell-in-range-format-as-percentage
' format for the percentage: https://stackoverflow.com/questions/42844778/vba-for-each-cell-in-range-format-as-percentage
    y = y + 1

    Else
    End If
Loop
End Sub
Sub Color()

Dim x, k As Integer
x = 1
k = Range("I1048576").End(xlUp).Row
' rownumber1: https://stackoverflow.com/questions/25056372/vba-range-row-count
For x = 2 To k
    If Cells(x, 10).Value > 0 Then
    Cells(x, 10).Interior.ColorIndex = 4
    ElseIf Cells(x, 10).Value < 0 Then
    Cells(x, 10).Interior.ColorIndex = 3
' color index: https://www.excel-easy.com/vba/examples/background-colors.html
    Else
    End If
Next x

End Sub
Sub MinMax()
Dim x, y, z As Integer
Dim Max, Min As Double
x = 1
Max = 0
Min = 1000000000

Do
    x = x + 1
    If Cells(x, 11).Value > Max Then
    Max = Cells(x, 11).Value
    y = x
    ElseIf Cells(x, 11).Value < Min Then
    Min = Cells(x, 11).Value
    z = x
    End If
Loop Until IsEmpty(Cells(x, 9).Value)
Cells(2, 16) = Max
Cells(2, 15) = Cells(y, 9).Value
Cells(3, 16) = Min
Cells(3, 15) = Cells(z, 9).Value
End Sub
Sub GreatVolume()
Dim x, y, k As Integer
Dim Max As Double
x = 1
Max = 0
k = Cells(Rows.Count, "i").End(xlUp).Row
For x = 2 To k
    If Cells(x, 12).Value > Max Then
    Max = Cells(x, 12).Value
    y = x
    End If
Next x

Cells(4, 16) = Max
Cells(4, 15) = Cells(y, 9).Value
End Sub
Sub Format()
Cells(2, 14).Value = "Greatest % Increase"
Cells(3, 14).Value = "Greatest % Decrease"
Cells(4, 14).Value = "Greatest Total Volume"
Cells(1, 15).Value = "Ticker"
Cells(1, 16).Value = "Value"
Range("P2", "P3").NumberFormat = "0.00%"
Range("P4").NumberFormat = "##0.0000E+0"
Range("N2:N5").Font.Bold = True
Range("O1:p1").Font.Bold = True
Columns("M:P").AutoFit
End Sub





