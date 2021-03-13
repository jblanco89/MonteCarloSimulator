Attribute VB_Name = "calculations"
Sub normalDist()
Dim N, Mean, Max, Min, sd, inputValue, I, ND, oRange As Range
Dim pMin, pMax As Double

Set ND = Hoja1.Range("I10")
Set I = Hoja1.Range("C5")
Set N = Hoja1.Range("C4")
Set Alpha = Hoja1.Range("C6")
Set inputValue = Hoja1.Range("C10")
Set sd = Hoja1.Range("G10")
Set Max = Hoja1.Range("F10")
Set Mean = Hoja1.Range("H10")
Set Min = Hoja1.Range("E10")

Row = 0
Row2 = 0
k = 1
For x = 1 To N.Value
    sd.Offset(Row, 0).Value = (Max.Offset(Row, 0) - inputValue.Offset(Row, 0)) / Application.WorksheetFunction.Norm_Inv(Alpha, 0, 1)
    If inputValue.Offset(Row, 0).Value = 0 Then
    Mean.Offset(Row, 0).Value = 0
    Else
    Mean.Offset(Row, 0).Value = Application.WorksheetFunction.Norm_Inv(Rnd, inputValue.Offset(Row, 0).Value, sd.Offset(Row, 0).Value)
    End If
    Row = Row + 1
    Next x
For j = 1 To I.Value
    ND.Offset(Row2, 0).Value = WorksheetFunction.Norm_Inv(Rnd, inputValue.Value, Max.Value)
'    ND.Offset(Row2, 0).Value = Rnd
    Row2 = Row2 + 1
    Next j
    
        Hoja1.Cells(9, 12) = "Count"
        Hoja1.Cells(10, 11).Formula = "=MIN(RC[-2]:R[" & I & "]C[-2])-MOD(MIN(RC[-2]:R[" & I & "]C[-2]),10)"
        Hoja1.Cells(20, 11).Formula = "=MAX(RC[-2]:R[" & I & "]C[-2])+10-MOD(MAX(RC[-2]:R[" & I & "]C[-2]),10)"
        pMin = Hoja1.Cells(10, 11)
        pMax = Hoja1.Cells(20, 11)

        For j = 11 To 19
        Hoja1.Cells(j, 11) = pMin + k * Round((pMax - pMin) / 10, 0)
        k = k + 1
        Next j
        
Set oRange = Hoja1.Range(Cells(10, 12), Cells(20, 12))
oRange.FormulaArray = "=FREQUENCY(RC[-3]:R[" & I & "]C[-3],RC[-1]:R[10]C[-1])"




End Sub



