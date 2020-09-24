Attribute VB_Name = "Module1"

Public Function Fgi(r As Integer, c As Integer, fg As Control) As Integer
   Fgi = c + fg.Cols * r
End Function
Public Sub MSHFlexGridEdit(MSHFlexGrid As Control, _
Edt As Control, KeyAscii As Integer)

Dim Tempy As Long

   Select Case KeyAscii
   Case 0 To 32
      Edt = MSHFlexGrid
      Edt.SelStart = 1000
   Case Else
      Edt = Chr(KeyAscii)
      Edt.SelStart = 1
   End Select
  Edt.Move MSHFlexGrid.Left + MSHFlexGrid.CellLeft, _
      MSHFlexGrid.Top + MSHFlexGrid.CellTop, _
      MSHFlexGrid.CellWidth - 7, _
      MSHFlexGrid.CellHeight - 7
 
   Edt.Visible = True
   Edt.SetFocus
End Sub
Public Sub EditKeyCode(MSHFlexGrid As Control, Edt As _
Control, KeyCode As Integer, Shift As Integer)
 Select Case KeyCode

   Case 27
      Edt.Visible = False
      MSHFlexGrid.SetFocus

   Case 13
      MSHFlexGrid.SetFocus

   Case 38
      MSHFlexGrid.SetFocus
      DoEvents
      If MSHFlexGrid.Row > MSHFlexGrid.FixedRows Then
         MSHFlexGrid.Row = MSHFlexGrid.Row - 1
      End If

   Case 40
      MSHFlexGrid.SetFocus
      DoEvents
      If MSHFlexGrid.Row < MSHFlexGrid.Rows - 1 Then
         MSHFlexGrid.Row = MSHFlexGrid.Row + 1
      End If
   End Select
End Sub



