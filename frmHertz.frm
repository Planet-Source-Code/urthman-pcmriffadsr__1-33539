VERSION 5.00
Begin VB.Form frmHertz 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Frequency"
   ClientHeight    =   1464
   ClientLeft      =   3456
   ClientTop       =   3288
   ClientWidth     =   3432
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   9.6
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1464
   ScaleWidth      =   3432
   Begin VB.CommandButton Command3 
      Caption         =   ">>"
      Height          =   375
      Index           =   1
      Left            =   2700
      TabIndex        =   4
      Top             =   300
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "<<"
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   300
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "OK"
      Height          =   375
      Left            =   1980
      TabIndex        =   2
      Top             =   960
      Width           =   1275
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Not"
      Height          =   375
      Left            =   180
      TabIndex        =   1
      Top             =   960
      Width           =   1275
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      MaxLength       =   5
      TabIndex        =   0
      Text            =   "00000"
      Top             =   300
      Width           =   1215
   End
End
Attribute VB_Name = "frmHertz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Code module and project created by "Urthman"
'   http://www.jsent.biz/urthman/
'   http://www.mp3.com/urthman/

Option Explicit

Dim FreqID%
Public Sub PrepMe(Which%)

    FreqID = Which
    Text1.Text = frmPcm.lblFreq(Which).Caption
    
End Sub


Private Sub Command1_Click()

    frmPcm.Enabled = True
    frmPcm.SetFocus
    Unload Me

End Sub


Private Sub Command2_Click()

    frmPcm.lblFreq(FreqID).Caption = Text1.Text
    frmPcm.Enabled = True
    frmPcm.SetFocus
    Unload Me

End Sub


Private Sub Command3_Click(Index As Integer)

    If IsNumeric(Text1.Text) Then
        Select Case Index
        Case 0: Text1.Text = (Text1.Text - 1)
        Case 1: Text1.Text = (Text1.Text + 1)
        End Select
    End If
    
End Sub

Private Sub Form_Load()

    Top = (Screen.Height - Height) / 2
    Left = (Screen.Width - Width) / 2

End Sub

Private Sub Text1_Change()

    If IsNumeric(Text1.Text) Then
        Command2.Enabled = (Val(Text1.Text) > 20) And (Val(Text1.Text) < 20000)
    Else
        Command2.Enabled = False
    End If

End Sub


Private Sub Text1_GotFocus()

    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.Text)

End Sub


Private Sub Text1_LostFocus()

    Text1.SelStart = 0
    Text1.SelLength = 0

End Sub


