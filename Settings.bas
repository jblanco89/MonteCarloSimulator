Attribute VB_Name = "Settings"
Sub clearContents()


lastrow = Hoja1.Cells(Hoja1.Rows.Count, "B").End(xlUp).Row

If lastrow > 9 Then

Hoja1.Range("B10:F" & lastrow).clearContents
Hoja1.Range("D10:D" & lastrow).Validation.Delete

Else

MsgBox "All data have been erased yet", vbExclamation

End If

End Sub



