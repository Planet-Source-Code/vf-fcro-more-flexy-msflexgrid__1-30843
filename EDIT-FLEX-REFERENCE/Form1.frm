VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Begin VB.Form Form1 
   Caption         =   "Extended MS Flex Grid by Vanja Fuckar,Zagreb,Croatia"
   ClientHeight    =   6495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11760
   LinkTopic       =   "Form1"
   ScaleHeight     =   6495
   ScaleWidth      =   11760
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00FFC0C0&
      Height          =   315
      Left            =   2760
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   2640
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFC0C0&
      Height          =   255
      Left            =   2760
      TabIndex        =   0
      Top             =   2160
      Visible         =   0   'False
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid MSH1 
      Height          =   4575
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   8070
      _Version        =   393216
      Rows            =   3
      Cols            =   3
      GridLines       =   2
      FormatString    =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Line Line2 
      X1              =   6000
      X2              =   6000
      Y1              =   5040
      Y2              =   5880
   End
   Begin VB.Shape Shape5 
      Height          =   855
      Left            =   1080
      Top             =   5040
      Width           =   9975
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      Caption         =   "EVENT-After Edit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   6120
      TabIndex        =   6
      Top             =   5160
      Width           =   4815
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "EVENT-Before Edit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   1200
      TabIndex        =   5
      Top             =   5160
      Width           =   4695
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   6120
      TabIndex        =   4
      Top             =   5520
      Width           =   4815
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "AUTHOR:VANJA FUCKAR,EMAIL:INGA@VIP.HR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   6000
      Width           =   5655
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   1200
      TabIndex        =   2
      Top             =   5520
      Width           =   4695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents FLX As EditableGrid
Attribute FLX.VB_VarHelpID = -1

Private Sub FLX_AfterEdit(row As Long, col As Long, NewValue As String)


Label1(1) = "ROW:" & Str(row) & ",COL: " & Str(col) & "  ,Value:" & NewValue

End Sub

Private Sub FLX_BeforeEdit(row As Long, col As Long, OldValue As String)
Dim datx(2) As String
Dim datxx(5) As String
If row = 1 And col = 1 Then

datx(0) = "a"
datx(1) = "b"
datx(2) = "c"
FLX.FillCombo datx
FLX.SetEditMethod = ListBox
FLX.SkipEdit = False

ElseIf row = 2 And col = 2 Then
datx(0) = "e"
datx(1) = "f"
datx(2) = "g"
FLX.FillCombo datx
FLX.SetEditMethod = ListBox
FLX.SkipEdit = False


ElseIf row = 2 And col = 3 Then
datxx(0) = "Huh!"
datxx(1) = "Oug!"
datxx(2) = "Hmm!"
datxx(3) = "Maybe?"
datxx(4) = "Brr."
datxx(5) = "MSX"
FLX.FillCombo datxx
FLX.SetEditMethod = ListBox
FLX.SkipEdit = False


ElseIf row = 3 And col = 3 Then
datx(0) = "1"
datx(1) = "2"
datx(2) = "3"
FLX.FillCombo datx
FLX.SetEditMethod = ListBox
FLX.SkipEdit = False


ElseIf row = 6 And col = 6 Then
datx(0) = "1"
datx(1) = "2"
datx(2) = "3"
FLX.FillCombo datx
FLX.SetEditMethod = ListBox
FLX.SkipEdit = False


ElseIf row = 4 And col = 4 Then
FLX.SkipEdit = True


ElseIf row = 4 And col = 2 Then
FLX.SkipEdit = True

Else
FLX.SetEditMethod = TextBox
FLX.SkipEdit = False
End If

Label1(0) = "ROW:" & Str(row) & ",COL: " & Str(col) & " ,Value: " & OldValue
End Sub



Private Sub Form_Load()
Set FLX = New EditableGrid


MSH1.Rows = 7
MSH1.Cols = 7

For u = 0 To MSH1.Cols - 1
MSH1.ColAlignment(u) = 1
Next u

Me.Top = (Screen.Height - Height) / 2
Me.Left = (Screen.Width - Width) / 2

FLX.SetControl MSH1, Text1, Combo1

MSH1.TextMatrix(0, 1) = "More"
MSH1.TextMatrix(0, 2) = "Flexy"
MSH1.TextMatrix(0, 3) = "MS"
MSH1.TextMatrix(0, 4) = "Grid."

MSH1.TextMatrix(1, 1) = "COMBO"
MSH1.TextMatrix(2, 2) = "COMBO"
MSH1.TextMatrix(3, 3) = "COMBO"
MSH1.TextMatrix(6, 6) = "COMBO"
MSH1.TextMatrix(2, 3) = "COMBO"
MSH1.TextMatrix(4, 4) = "Skip Edit"
MSH1.TextMatrix(4, 2) = "Skip Edit"
End Sub

