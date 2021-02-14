VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Clear_Click()
Call clearContents

End Sub


Private Sub worksheet_change(ByVal target As Range)
Dim N As Range
Dim I As Range
Dim Alpha As Range
Dim styleRange As Range


Set N = Range("C4")
Set I = Range("C5")
Set Alpha = Range("C6")


If Not Application.Intersect(N, Range(target.Address)) _
    Is Nothing Then
    lastrow = Hoja1.Cells(Hoja1.Rows.Count, "B").End(xlUp).Row + 1
    Hoja1.Range("B10:F" & lastrow).Font.Name = "Ubuntu"
    Hoja1.Range("B10:F" & lastrow).Font.Size = 12
    Hoja1.Range("B10:F" & lastrow).HorizontalAlignment = xlCenter
    Hoja1.Range("B10:B" & lastrow).clearContents
    Hoja1.Range("D10:D" & lastrow).Validation.Delete
    Row = 0
    For x = 1 To N.Value
    Hoja1.Cells(10, "B").Offset(Row, 0).Value = x
    Hoja1.Cells(10, "D").Offset(Row, 0).Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
    Formula1:="Normal, Binomial, Uniform, Triangular, Logaritm"
    Row = Row + 1
    Next x
End If

    If I.Value > 10000 Or I.Value = 0 Then
    MsgBox "Num Iteration should be between 1 and 10,000, please check", vbCritical, Warning
    I.Value = 1

End If
    
If Not Application.Intersect(Alpha, Range(target.Address)) _
    Is Nothing Then
    MsgBox "Confidence Selected succesfully"
End If


End Sub


