VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmAdsr 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ADSR Envelope"
   ClientHeight    =   2340
   ClientLeft      =   2064
   ClientTop       =   4176
   ClientWidth     =   7644
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   10.2
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   7644
   Begin VB.CommandButton Command1 
      Caption         =   "Accept"
      Height          =   375
      Left            =   6060
      TabIndex        =   12
      Top             =   1860
      Width           =   1395
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   180
      TabIndex        =   11
      Top             =   1860
      Width           =   1395
   End
   Begin MSComctlLib.Slider sldAttack 
      Height          =   312
      Left            =   720
      TabIndex        =   6
      Top             =   1200
      Width           =   1692
      _ExtentX        =   2985
      _ExtentY        =   550
      _Version        =   393216
      Max             =   130
      SelStart        =   1
      TickFrequency   =   10
      Value           =   1
   End
   Begin MSComctlLib.Slider sldPeak 
      Height          =   1692
      Left            =   2700
      TabIndex        =   7
      Top             =   540
      Width           =   312
      _ExtentX        =   550
      _ExtentY        =   2985
      _Version        =   393216
      Orientation     =   1
      Max             =   99
      SelStart        =   1
      TickFrequency   =   10
      Value           =   1
   End
   Begin MSComctlLib.Slider sldDecay 
      Height          =   312
      Left            =   3240
      TabIndex        =   8
      Top             =   1200
      Width           =   1692
      _ExtentX        =   2985
      _ExtentY        =   550
      _Version        =   393216
      Max             =   130
      SelStart        =   1
      TickFrequency   =   10
      Value           =   1
   End
   Begin MSComctlLib.Slider sldSust 
      Height          =   1692
      Left            =   5280
      TabIndex        =   9
      Top             =   540
      Width           =   312
      _ExtentX        =   550
      _ExtentY        =   2985
      _Version        =   393216
      Orientation     =   1
      Max             =   99
      SelStart        =   1
      TickFrequency   =   10
      Value           =   1
   End
   Begin MSComctlLib.Slider sldRel 
      Height          =   312
      Left            =   5760
      TabIndex        =   10
      Top             =   1200
      Width           =   1692
      _ExtentX        =   2985
      _ExtentY        =   550
      _Version        =   393216
      Max             =   130
      SelStart        =   1
      TickFrequency   =   10
      Value           =   1
   End
   Begin VB.Label lblFreq 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   312
      Left            =   240
      TabIndex        =   5
      ToolTipText     =   "Frequency"
      Top             =   120
      Width           =   1152
   End
   Begin VB.Label lblPeak 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   312
      Left            =   2220
      TabIndex        =   4
      ToolTipText     =   "Peak Level (db)"
      Top             =   120
      Width           =   1152
   End
   Begin VB.Label lblAttack 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   312
      Left            =   1020
      TabIndex        =   3
      ToolTipText     =   "Attack (Rise) Time"
      Top             =   780
      Width           =   1152
   End
   Begin VB.Label lblDecay 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   312
      Left            =   3480
      TabIndex        =   2
      ToolTipText     =   "Decay Time"
      Top             =   780
      Width           =   1152
   End
   Begin VB.Label lblSust 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   312
      Left            =   4800
      TabIndex        =   1
      ToolTipText     =   "Sustain Level (db)"
      Top             =   120
      Width           =   1152
   End
   Begin VB.Label lblRel 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   312
      Left            =   6000
      TabIndex        =   0
      ToolTipText     =   "Release Time"
      Top             =   780
      Width           =   1152
   End
End
Attribute VB_Name = "frmAdsr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ThisGuy%
Public Sub SetMeUp(Which%)

    ThisGuy = Which

    With frmPcm

    lblFreq.Caption = .lblFreq(Which).Caption
    
    sldAttack.Value = (.lblAttack(Which).Caption / 10)
    sldAttack_Click
    sldPeak.Value = (0 - (.lblPeak(Which).Caption * 3))
    sldPeak_Click
    sldDecay.Value = (.lblDecay(Which).Caption / 10)
    sldDecay_Click
    sldSust.Value = (0 - (.lblSustain(Which).Caption * 3))
    sldSust_Click
    sldRel.Value = (.lblRelease(Which).Caption / 10)
    sldRel_Click

    End With

End Sub

Private Sub Command1_Click()

    With frmPcm
    
    .lblAttack(ThisGuy).Caption = lblAttack.Caption
    .lblPeak(ThisGuy).Caption = lblPeak.Caption
    .lblDecay(ThisGuy).Caption = lblDecay.Caption
    .lblSustain(ThisGuy).Caption = lblSust.Caption
    .lblRelease(ThisGuy).Caption = lblRel.Caption

    End With

    frmPcm.Enabled = True
    frmPcm.SetFocus
    Unload Me

End Sub


Private Sub Command2_Click()

    frmPcm.Enabled = True
    frmPcm.SetFocus
    Unload Me

End Sub

Private Sub Form_Load()

    Top = (Screen.Height - Height) / 2
    Left = (Screen.Width - Width) / 2

End Sub


Private Sub sldAttack_Change()

    lblAttack.Caption = (sldAttack.Value * 10)
    
End Sub

Private Sub sldAttack_Click()

    lblAttack.Caption = (sldAttack.Value * 10)

End Sub


Private Sub sldDecay_Change()

    lblDecay.Caption = (sldDecay.Value * 10)

End Sub

Private Sub sldDecay_Click()

    lblDecay.Caption = (sldDecay.Value * 10)

End Sub

Private Sub sldPeak_Change()

    lblPeak.Caption = Format((0 - (sldPeak.Value / 3)), "#0.00")

End Sub

Private Sub sldPeak_Click()

    lblPeak.Caption = Format((0 - (sldPeak.Value / 3)), "#0.00")

End Sub


Private Sub sldRel_Change()

    lblRel.Caption = (sldRel.Value * 10)

End Sub

Private Sub sldRel_Click()

    lblRel.Caption = (sldRel.Value * 10)

End Sub

Private Sub sldSust_Change()

    lblSust.Caption = Format((0 - (sldSust.Value / 3)), "#0.00")

End Sub

Private Sub sldSust_Click()

    lblSust.Caption = Format((0 - (sldSust.Value / 3)), "#0.00")

End Sub


