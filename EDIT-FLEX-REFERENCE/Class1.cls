VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EditableGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private WithEvents txtedit As TextBox
Attribute txtedit.VB_VarHelpID = -1
Private WithEvents fg2 As MSFlexGrid
Attribute fg2.VB_VarHelpID = -1
Private WithEvents combox As ComboBox
Attribute combox.VB_VarHelpID = -1

Private Delay As Boolean
Private Datax() As String

Public Enum MethodType
 TextBox = 0
 ListBox = 1
End Enum

Private Method As MethodType
Event BeforeEdit(Row As Long, Col As Long, OldValue As String)
Event AfterEdit(Row As Long, Col As Long, NewValue As String)
Public Property Let SetEditMethod(CellMethod As MethodType)
Method = CellMethod
End Property
Public Property Let SkipEdit(allow As Boolean)
Delay = allow
End Property
Private Sub Class_Terminate()
Set txtedit = Nothing
Set fg2 = Nothing
Set combox = Nothing
Erase Datax
End Sub
Public Sub FillCombo(data() As String)
Erase Datax
Datax = data
End Sub
Private Sub combox_Click()
ProcCombo
End Sub

Private Sub combox_DblClick()
ProcCombo
End Sub

Private Sub fg2_Click()
ProcCombo
End Sub

Private Sub Fg2_KeyPress(KeyAscii As Integer)
BeforePROC
ProcCombo
If Delay = True Then Exit Sub
If Method = TextBox Then
MSHFlexGridEdit fg2, txtedit, KeyAscii
Else
ComboProp
End If
End Sub
Private Sub Fg2_DblClick()
BeforePROC
ProcCombo
If Delay = True Then Exit Sub
If Method = TextBox Then
MSHFlexGridEdit fg2, txtedit, 32
Else
ComboProp
End If
End Sub
Private Sub Fg2_LeaveCell()
ProcCombo
AfterPROC
End Sub
Private Sub Fg2_GotFocus()
ProcCombo
AfterPROC
End Sub

Private Sub fg2_Scroll()
ProcCombo
AfterPROC
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
If KeyAscii = Asc(vbCr) Then KeyAscii = 0
End Sub
Private Sub txtEdit_KeyDown(KeyCode As Integer, _
Shift As Integer)
EditKeyCode fg2, txtedit, KeyCode, Shift
End Sub
Public Sub SetControl(FlxGrid As Object, Optional TxtBox As Object, Optional CmBox As Object)
On Error GoTo ext
Set fg2 = FlxGrid
Set txtedit = TxtBox
Set combox = CmBox
txtedit.Visible = False
CmBox.Visible = False
Align
Fnt
fg2.RowHeightMin = combox.Height + 8
ext:
If Err <> 0 Then On Error GoTo 0
End Sub
Private Function CalcAlign() As Long
Select Case fg2.ColAlignment(fg2.Col)
Case Is = 0, 1, 2, 9
CalcAlign = 0
Case Is = 3, 4, 5
CalcAlign = 2
Case Is = 6, 7, 8
CalcAlign = 1
End Select
End Function

Private Sub Align()
Dim ClcA As Long
ClcA = CalcAlign
txtedit.Alignment = ClcA
End Sub
Private Sub Fnt()
txtedit.FontBold = fg2.CellFontBold
txtedit.FontItalic = fg2.CellFontItalic
txtedit.FontName = fg2.CellFontName
txtedit.FontSize = fg2.CellFontSize
txtedit.FontStrikethru = fg2.CellFontStrikeThrough
txtedit.FontUnderline = fg2.CellFontUnderline
txtedit.Font.Charset = fg2.Font.Charset
combox.FontBold = fg2.CellFontBold
combox.FontItalic = fg2.CellFontItalic
combox.FontName = fg2.CellFontName
combox.FontSize = fg2.CellFontSize
combox.FontStrikethru = fg2.CellFontStrikeThrough
combox.FontUnderline = fg2.CellFontUnderline
combox.Font.Charset = fg2.Font.Charset
End Sub
Private Sub BeforePROC()
RaiseEvent BeforeEdit(fg2.Row, fg2.Col, fg2)
End Sub
Private Sub AfterPROC()
If txtedit.Visible = False Then Exit Sub
fg2 = txtedit
txtedit.Visible = False
RaiseEvent AfterEdit(fg2.Row, fg2.Col, fg2)
End Sub
Private Sub txtedit_LostFocus()
AfterPROC
End Sub

Private Sub ProcCombo()
If combox.Visible = False Then Exit Sub
If combox.ListIndex <> -1 Then fg2 = combox
RaiseEvent AfterEdit(fg2.Row, fg2.Col, fg2)
combox.Visible = False
End Sub
Private Sub ComboProp()
On Error GoTo ext
combox.Left = fg2.Left + fg2.CellLeft
combox.Top = fg2.Top + fg2.CellTop
combox.Width = fg2.CellWidth - 7
combox.Clear
Dim Xz As Long
Xz = -1
For u = 0 To UBound(Datax)
combox.AddItem Datax(u)
If fg2 = Datax(u) Then Xz = u
Next u
combox.ListIndex = Xz
combox.Visible = True
ext:
If Err <> 0 Then On Error GoTo 0
End Sub


