VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmPcm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PCM"
   ClientHeight    =   5592
   ClientLeft      =   660
   ClientTop       =   1692
   ClientWidth     =   9732
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
   Moveable        =   0   'False
   ScaleHeight     =   5592
   ScaleWidth      =   9732
   Begin VB.Frame Frame1 
      Caption         =   "Wave Maker"
      Height          =   5055
      Index           =   1
      Left            =   60
      TabIndex        =   2
      Top             =   0
      Width           =   9615
      Begin VB.CheckBox chkLFO3 
         BackColor       =   &H00FFC0FF&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   12
         Left            =   9180
         Style           =   1  'Graphical
         TabIndex        =   107
         Top             =   4620
         Width           =   315
      End
      Begin VB.CheckBox chkLFO2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   12
         Left            =   8820
         Style           =   1  'Graphical
         TabIndex        =   106
         Top             =   4620
         Width           =   315
      End
      Begin VB.CheckBox chkLFO1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   12
         Left            =   8460
         Style           =   1  'Graphical
         TabIndex        =   105
         Top             =   4620
         Width           =   315
      End
      Begin VB.CheckBox chkLFO3 
         BackColor       =   &H00FFC0FF&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   11
         Left            =   9180
         Style           =   1  'Graphical
         TabIndex        =   104
         Top             =   4260
         Width           =   315
      End
      Begin VB.CheckBox chkLFO2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   11
         Left            =   8820
         Style           =   1  'Graphical
         TabIndex        =   103
         Top             =   4260
         Width           =   315
      End
      Begin VB.CheckBox chkLFO1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   11
         Left            =   8460
         Style           =   1  'Graphical
         TabIndex        =   102
         Top             =   4260
         Width           =   315
      End
      Begin VB.CheckBox chkLFO3 
         BackColor       =   &H00FFC0FF&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   10
         Left            =   9180
         Style           =   1  'Graphical
         TabIndex        =   101
         Top             =   3900
         Width           =   315
      End
      Begin VB.CheckBox chkLFO2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   10
         Left            =   8820
         Style           =   1  'Graphical
         TabIndex        =   100
         Top             =   3900
         Width           =   315
      End
      Begin VB.CheckBox chkLFO1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   10
         Left            =   8460
         Style           =   1  'Graphical
         TabIndex        =   99
         Top             =   3900
         Width           =   315
      End
      Begin VB.CheckBox chkLFO3 
         BackColor       =   &H00FFC0FF&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   9
         Left            =   9180
         Style           =   1  'Graphical
         TabIndex        =   98
         Top             =   3540
         Width           =   315
      End
      Begin VB.CheckBox chkLFO2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   9
         Left            =   8820
         Style           =   1  'Graphical
         TabIndex        =   97
         Top             =   3540
         Width           =   315
      End
      Begin VB.CheckBox chkLFO1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   9
         Left            =   8460
         Style           =   1  'Graphical
         TabIndex        =   96
         Top             =   3540
         Width           =   315
      End
      Begin VB.CheckBox chkLFO3 
         BackColor       =   &H00FFC0FF&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   8
         Left            =   9180
         Style           =   1  'Graphical
         TabIndex        =   95
         Top             =   3180
         Width           =   315
      End
      Begin VB.CheckBox chkLFO2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   8
         Left            =   8820
         Style           =   1  'Graphical
         TabIndex        =   94
         Top             =   3180
         Width           =   315
      End
      Begin VB.CheckBox chkLFO1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   8
         Left            =   8460
         Style           =   1  'Graphical
         TabIndex        =   93
         Top             =   3180
         Width           =   315
      End
      Begin VB.CheckBox chkLFO3 
         BackColor       =   &H00FFC0FF&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   7
         Left            =   9180
         Style           =   1  'Graphical
         TabIndex        =   92
         Top             =   2820
         Width           =   315
      End
      Begin VB.CheckBox chkLFO2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   7
         Left            =   8820
         Style           =   1  'Graphical
         TabIndex        =   91
         Top             =   2820
         Width           =   315
      End
      Begin VB.CheckBox chkLFO1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   7
         Left            =   8460
         Style           =   1  'Graphical
         TabIndex        =   90
         Top             =   2820
         Width           =   315
      End
      Begin VB.CheckBox chkLFO3 
         BackColor       =   &H00FFC0FF&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   6
         Left            =   9180
         Style           =   1  'Graphical
         TabIndex        =   89
         Top             =   2460
         Width           =   315
      End
      Begin VB.CheckBox chkLFO2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   6
         Left            =   8820
         Style           =   1  'Graphical
         TabIndex        =   88
         Top             =   2460
         Width           =   315
      End
      Begin VB.CheckBox chkLFO1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   6
         Left            =   8460
         Style           =   1  'Graphical
         TabIndex        =   87
         Top             =   2460
         Width           =   315
      End
      Begin VB.CheckBox chkLFO3 
         BackColor       =   &H00FFC0FF&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         Left            =   9180
         Style           =   1  'Graphical
         TabIndex        =   86
         Top             =   2100
         Width           =   315
      End
      Begin VB.CheckBox chkLFO2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         Left            =   8820
         Style           =   1  'Graphical
         TabIndex        =   85
         Top             =   2100
         Width           =   315
      End
      Begin VB.CheckBox chkLFO1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         Left            =   8460
         Style           =   1  'Graphical
         TabIndex        =   84
         Top             =   2100
         Width           =   315
      End
      Begin VB.CheckBox chkLFO3 
         BackColor       =   &H00FFC0FF&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   9180
         Style           =   1  'Graphical
         TabIndex        =   83
         Top             =   1740
         Width           =   315
      End
      Begin VB.CheckBox chkLFO2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   8820
         Style           =   1  'Graphical
         TabIndex        =   82
         Top             =   1740
         Width           =   315
      End
      Begin VB.CheckBox chkLFO1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   8460
         Style           =   1  'Graphical
         TabIndex        =   81
         Top             =   1740
         Width           =   315
      End
      Begin VB.CheckBox chkLFO3 
         BackColor       =   &H00FFC0FF&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   9180
         Style           =   1  'Graphical
         TabIndex        =   80
         Top             =   1380
         Width           =   315
      End
      Begin VB.CheckBox chkLFO2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   8820
         Style           =   1  'Graphical
         TabIndex        =   79
         Top             =   1380
         Width           =   315
      End
      Begin VB.CheckBox chkLFO1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   8460
         Style           =   1  'Graphical
         TabIndex        =   78
         Top             =   1380
         Width           =   315
      End
      Begin VB.CheckBox chkLFO3 
         BackColor       =   &H00FFC0FF&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   9180
         Style           =   1  'Graphical
         TabIndex        =   77
         Top             =   1020
         Width           =   315
      End
      Begin VB.CheckBox chkLFO2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   8820
         Style           =   1  'Graphical
         TabIndex        =   76
         Top             =   1020
         Width           =   315
      End
      Begin VB.CheckBox chkLFO1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   8460
         Style           =   1  'Graphical
         TabIndex        =   75
         Top             =   1020
         Width           =   315
      End
      Begin VB.CheckBox chkLFO3 
         BackColor       =   &H00FFC0FF&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   9180
         Style           =   1  'Graphical
         TabIndex        =   74
         Top             =   660
         Width           =   315
      End
      Begin VB.CheckBox chkLFO2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   8820
         Style           =   1  'Graphical
         TabIndex        =   73
         Top             =   660
         Width           =   315
      End
      Begin VB.CheckBox chkLFO1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   8460
         Style           =   1  'Graphical
         TabIndex        =   72
         Top             =   660
         Width           =   315
      End
      Begin VB.CheckBox chkLFO3 
         BackColor       =   &H00FFC0FF&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   9180
         Style           =   1  'Graphical
         TabIndex        =   71
         Top             =   300
         Width           =   315
      End
      Begin VB.CheckBox chkLFO2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   8820
         Style           =   1  'Graphical
         TabIndex        =   70
         Top             =   300
         Width           =   315
      End
      Begin VB.CheckBox chkLFO1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   8460
         Style           =   1  'Graphical
         TabIndex        =   69
         Top             =   300
         Width           =   315
      End
      Begin VB.CheckBox chkChan 
         BackColor       =   &H00C0C0FF&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   12
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   4620
         Width           =   315
      End
      Begin VB.CheckBox chkChan 
         BackColor       =   &H00C0C0FF&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   11
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   4260
         Width           =   315
      End
      Begin VB.CheckBox chkChan 
         BackColor       =   &H00C0C0FF&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   10
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   3900
         Width           =   315
      End
      Begin VB.CheckBox chkChan 
         BackColor       =   &H00C0C0FF&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   9
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   3540
         Width           =   315
      End
      Begin VB.CheckBox chkChan 
         BackColor       =   &H00C0C0FF&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   8
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   3180
         Width           =   315
      End
      Begin VB.CheckBox chkChan 
         BackColor       =   &H00C0C0FF&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   7
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   2820
         Width           =   315
      End
      Begin VB.CheckBox chkChan 
         BackColor       =   &H00C0C0FF&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   6
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   2460
         Width           =   315
      End
      Begin VB.CheckBox chkChan 
         BackColor       =   &H00C0C0FF&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   2100
         Width           =   315
      End
      Begin VB.CheckBox chkChan 
         BackColor       =   &H00C0C0FF&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   1740
         Width           =   315
      End
      Begin VB.CheckBox chkChan 
         BackColor       =   &H00C0C0FF&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   1380
         Width           =   315
      End
      Begin VB.CheckBox chkChan 
         BackColor       =   &H00C0C0FF&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   1020
         Width           =   315
      End
      Begin VB.CheckBox chkChan 
         BackColor       =   &H00C0C0FF&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   660
         Width           =   315
      End
      Begin VB.CheckBox chkChan 
         BackColor       =   &H00C0C0FF&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   300
         Width           =   315
      End
      Begin VB.Label lblRelease 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   12
         Left            =   7020
         TabIndex        =   176
         ToolTipText     =   "Release Time"
         Top             =   4620
         Width           =   1155
      End
      Begin VB.Label lblSustain 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   12
         Left            =   5760
         TabIndex        =   175
         ToolTipText     =   "Sustain Level (db)"
         Top             =   4620
         Width           =   1155
      End
      Begin VB.Label lblDecay 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   12
         Left            =   4500
         TabIndex        =   174
         ToolTipText     =   "Decay Time"
         Top             =   4620
         Width           =   1155
      End
      Begin VB.Label lblAttack 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   12
         Left            =   1980
         TabIndex        =   173
         ToolTipText     =   "Attack (Rise) Time"
         Top             =   4620
         Width           =   1155
      End
      Begin VB.Label lblRelease 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   11
         Left            =   7020
         TabIndex        =   172
         ToolTipText     =   "Release Time"
         Top             =   4260
         Width           =   1155
      End
      Begin VB.Label lblSustain 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   11
         Left            =   5760
         TabIndex        =   171
         ToolTipText     =   "Sustain Level (db)"
         Top             =   4260
         Width           =   1155
      End
      Begin VB.Label lblDecay 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   11
         Left            =   4500
         TabIndex        =   170
         ToolTipText     =   "Decay Time"
         Top             =   4260
         Width           =   1155
      End
      Begin VB.Label lblAttack 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   11
         Left            =   1980
         TabIndex        =   169
         ToolTipText     =   "Attack (Rise) Time"
         Top             =   4260
         Width           =   1155
      End
      Begin VB.Label lblRelease 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   10
         Left            =   7020
         TabIndex        =   168
         ToolTipText     =   "Release Time"
         Top             =   3900
         Width           =   1155
      End
      Begin VB.Label lblSustain 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   10
         Left            =   5760
         TabIndex        =   167
         ToolTipText     =   "Sustain Level (db)"
         Top             =   3900
         Width           =   1155
      End
      Begin VB.Label lblDecay 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   10
         Left            =   4500
         TabIndex        =   166
         ToolTipText     =   "Decay Time"
         Top             =   3900
         Width           =   1155
      End
      Begin VB.Label lblAttack 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   10
         Left            =   1980
         TabIndex        =   165
         ToolTipText     =   "Attack (Rise) Time"
         Top             =   3900
         Width           =   1155
      End
      Begin VB.Label lblRelease 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   9
         Left            =   7020
         TabIndex        =   164
         ToolTipText     =   "Release Time"
         Top             =   3540
         Width           =   1155
      End
      Begin VB.Label lblSustain 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   9
         Left            =   5760
         TabIndex        =   163
         ToolTipText     =   "Sustain Level (db)"
         Top             =   3540
         Width           =   1155
      End
      Begin VB.Label lblDecay 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   9
         Left            =   4500
         TabIndex        =   162
         ToolTipText     =   "Decay Time"
         Top             =   3540
         Width           =   1155
      End
      Begin VB.Label lblAttack 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   9
         Left            =   1980
         TabIndex        =   161
         ToolTipText     =   "Attack (Rise) Time"
         Top             =   3540
         Width           =   1155
      End
      Begin VB.Label lblRelease 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   8
         Left            =   7020
         TabIndex        =   160
         ToolTipText     =   "Release Time"
         Top             =   3180
         Width           =   1155
      End
      Begin VB.Label lblSustain 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   8
         Left            =   5760
         TabIndex        =   159
         ToolTipText     =   "Sustain Level (db)"
         Top             =   3180
         Width           =   1155
      End
      Begin VB.Label lblDecay 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   8
         Left            =   4500
         TabIndex        =   158
         ToolTipText     =   "Decay Time"
         Top             =   3180
         Width           =   1155
      End
      Begin VB.Label lblAttack 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   8
         Left            =   1980
         TabIndex        =   157
         ToolTipText     =   "Attack (Rise) Time"
         Top             =   3180
         Width           =   1155
      End
      Begin VB.Label lblRelease 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   7
         Left            =   7020
         TabIndex        =   156
         ToolTipText     =   "Release Time"
         Top             =   2820
         Width           =   1155
      End
      Begin VB.Label lblSustain 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   7
         Left            =   5760
         TabIndex        =   155
         ToolTipText     =   "Sustain Level (db)"
         Top             =   2820
         Width           =   1155
      End
      Begin VB.Label lblDecay 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   7
         Left            =   4500
         TabIndex        =   154
         ToolTipText     =   "Decay Time"
         Top             =   2820
         Width           =   1155
      End
      Begin VB.Label lblAttack 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   7
         Left            =   1980
         TabIndex        =   153
         ToolTipText     =   "Attack (Rise) Time"
         Top             =   2820
         Width           =   1155
      End
      Begin VB.Label lblRelease 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   6
         Left            =   7020
         TabIndex        =   152
         ToolTipText     =   "Release Time"
         Top             =   2460
         Width           =   1155
      End
      Begin VB.Label lblSustain 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   6
         Left            =   5760
         TabIndex        =   151
         ToolTipText     =   "Sustain Level (db)"
         Top             =   2460
         Width           =   1155
      End
      Begin VB.Label lblDecay 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   6
         Left            =   4500
         TabIndex        =   150
         ToolTipText     =   "Decay Time"
         Top             =   2460
         Width           =   1155
      End
      Begin VB.Label lblAttack 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   6
         Left            =   1980
         TabIndex        =   149
         ToolTipText     =   "Attack (Rise) Time"
         Top             =   2460
         Width           =   1155
      End
      Begin VB.Label lblRelease 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   5
         Left            =   7020
         TabIndex        =   148
         ToolTipText     =   "Release Time"
         Top             =   2100
         Width           =   1155
      End
      Begin VB.Label lblSustain 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   5
         Left            =   5760
         TabIndex        =   147
         ToolTipText     =   "Sustain Level (db)"
         Top             =   2100
         Width           =   1155
      End
      Begin VB.Label lblDecay 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   5
         Left            =   4500
         TabIndex        =   146
         ToolTipText     =   "Decay Time"
         Top             =   2100
         Width           =   1155
      End
      Begin VB.Label lblAttack 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   5
         Left            =   1980
         TabIndex        =   145
         ToolTipText     =   "Attack (Rise) Time"
         Top             =   2100
         Width           =   1155
      End
      Begin VB.Label lblRelease 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   4
         Left            =   7020
         TabIndex        =   144
         ToolTipText     =   "Release Time"
         Top             =   1740
         Width           =   1155
      End
      Begin VB.Label lblSustain 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   4
         Left            =   5760
         TabIndex        =   143
         ToolTipText     =   "Sustain Level (db)"
         Top             =   1740
         Width           =   1155
      End
      Begin VB.Label lblDecay 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   4
         Left            =   4500
         TabIndex        =   142
         ToolTipText     =   "Decay Time"
         Top             =   1740
         Width           =   1155
      End
      Begin VB.Label lblAttack 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   4
         Left            =   1980
         TabIndex        =   141
         ToolTipText     =   "Attack (Rise) Time"
         Top             =   1740
         Width           =   1155
      End
      Begin VB.Label lblRelease 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   3
         Left            =   7020
         TabIndex        =   140
         ToolTipText     =   "Release Time"
         Top             =   1380
         Width           =   1155
      End
      Begin VB.Label lblSustain 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   3
         Left            =   5760
         TabIndex        =   139
         ToolTipText     =   "Sustain Level (db)"
         Top             =   1380
         Width           =   1155
      End
      Begin VB.Label lblDecay 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   3
         Left            =   4500
         TabIndex        =   138
         ToolTipText     =   "Decay Time"
         Top             =   1380
         Width           =   1155
      End
      Begin VB.Label lblAttack 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   3
         Left            =   1980
         TabIndex        =   137
         ToolTipText     =   "Attack (Rise) Time"
         Top             =   1380
         Width           =   1155
      End
      Begin VB.Label lblRelease 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   2
         Left            =   7020
         TabIndex        =   136
         ToolTipText     =   "Release Time"
         Top             =   1020
         Width           =   1155
      End
      Begin VB.Label lblSustain 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   2
         Left            =   5760
         TabIndex        =   135
         ToolTipText     =   "Sustain Level (db)"
         Top             =   1020
         Width           =   1155
      End
      Begin VB.Label lblDecay 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   2
         Left            =   4500
         TabIndex        =   134
         ToolTipText     =   "Decay Time"
         Top             =   1020
         Width           =   1155
      End
      Begin VB.Label lblAttack 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   2
         Left            =   1980
         TabIndex        =   133
         ToolTipText     =   "Attack (Rise) Time"
         Top             =   1020
         Width           =   1155
      End
      Begin VB.Label lblRelease 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   1
         Left            =   7020
         TabIndex        =   132
         ToolTipText     =   "Release Time"
         Top             =   660
         Width           =   1155
      End
      Begin VB.Label lblSustain 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   1
         Left            =   5760
         TabIndex        =   131
         ToolTipText     =   "Sustain Level (db)"
         Top             =   660
         Width           =   1155
      End
      Begin VB.Label lblDecay 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   1
         Left            =   4500
         TabIndex        =   130
         ToolTipText     =   "Decay Time"
         Top             =   660
         Width           =   1155
      End
      Begin VB.Label lblAttack 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   1
         Left            =   1980
         TabIndex        =   129
         ToolTipText     =   "Attack (Rise) Time"
         Top             =   660
         Width           =   1155
      End
      Begin VB.Label lblRelease 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   0
         Left            =   7020
         TabIndex        =   128
         ToolTipText     =   "Release Time"
         Top             =   300
         Width           =   1155
      End
      Begin VB.Label lblSustain 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   0
         Left            =   5760
         TabIndex        =   127
         ToolTipText     =   "Sustain Level (db)"
         Top             =   300
         Width           =   1155
      End
      Begin VB.Label lblDecay 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   0
         Left            =   4500
         TabIndex        =   126
         ToolTipText     =   "Decay Time"
         Top             =   300
         Width           =   1155
      End
      Begin VB.Label lblAttack 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   0
         Left            =   1980
         TabIndex        =   125
         ToolTipText     =   "Attack (Rise) Time"
         Top             =   300
         Width           =   1155
      End
      Begin VB.Label lblPeak 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   12
         Left            =   3240
         TabIndex        =   30
         ToolTipText     =   "Peak Level (db)"
         Top             =   4620
         Width           =   1155
      End
      Begin VB.Label lblFreq 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   12
         Left            =   480
         TabIndex        =   29
         ToolTipText     =   "Frequency"
         Top             =   4620
         Width           =   1155
      End
      Begin VB.Label lblPeak 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   11
         Left            =   3240
         TabIndex        =   28
         ToolTipText     =   "Peak Level (db)"
         Top             =   4260
         Width           =   1155
      End
      Begin VB.Label lblFreq 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   11
         Left            =   480
         TabIndex        =   27
         ToolTipText     =   "Frequency"
         Top             =   4260
         Width           =   1155
      End
      Begin VB.Label lblPeak 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   10
         Left            =   3240
         TabIndex        =   26
         ToolTipText     =   "Peak Level (db)"
         Top             =   3900
         Width           =   1155
      End
      Begin VB.Label lblFreq 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   10
         Left            =   480
         TabIndex        =   25
         ToolTipText     =   "Frequency"
         Top             =   3900
         Width           =   1155
      End
      Begin VB.Label lblPeak 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   9
         Left            =   3240
         TabIndex        =   24
         ToolTipText     =   "Peak Level (db)"
         Top             =   3540
         Width           =   1155
      End
      Begin VB.Label lblFreq 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   9
         Left            =   480
         TabIndex        =   23
         ToolTipText     =   "Frequency"
         Top             =   3540
         Width           =   1155
      End
      Begin VB.Label lblPeak 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   8
         Left            =   3240
         TabIndex        =   22
         ToolTipText     =   "Peak Level (db)"
         Top             =   3180
         Width           =   1155
      End
      Begin VB.Label lblFreq 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   8
         Left            =   480
         TabIndex        =   21
         ToolTipText     =   "Frequency"
         Top             =   3180
         Width           =   1155
      End
      Begin VB.Label lblPeak 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   7
         Left            =   3240
         TabIndex        =   20
         ToolTipText     =   "Peak Level (db)"
         Top             =   2820
         Width           =   1155
      End
      Begin VB.Label lblFreq 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   7
         Left            =   480
         TabIndex        =   19
         ToolTipText     =   "Frequency"
         Top             =   2820
         Width           =   1155
      End
      Begin VB.Label lblPeak 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   6
         Left            =   3240
         TabIndex        =   18
         ToolTipText     =   "Peak Level (db)"
         Top             =   2460
         Width           =   1155
      End
      Begin VB.Label lblFreq 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   6
         Left            =   480
         TabIndex        =   17
         ToolTipText     =   "Frequency"
         Top             =   2460
         Width           =   1155
      End
      Begin VB.Label lblPeak 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   5
         Left            =   3240
         TabIndex        =   16
         ToolTipText     =   "Peak Level (db)"
         Top             =   2100
         Width           =   1155
      End
      Begin VB.Label lblFreq 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   5
         Left            =   480
         TabIndex        =   15
         ToolTipText     =   "Frequency"
         Top             =   2100
         Width           =   1155
      End
      Begin VB.Label lblPeak 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   4
         Left            =   3240
         TabIndex        =   14
         ToolTipText     =   "Peak Level (db)"
         Top             =   1740
         Width           =   1155
      End
      Begin VB.Label lblFreq 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   4
         Left            =   480
         TabIndex        =   13
         ToolTipText     =   "Frequency"
         Top             =   1740
         Width           =   1155
      End
      Begin VB.Label lblPeak 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   3
         Left            =   3240
         TabIndex        =   12
         ToolTipText     =   "Peak Level (db)"
         Top             =   1380
         Width           =   1155
      End
      Begin VB.Label lblFreq 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   3
         Left            =   480
         TabIndex        =   11
         ToolTipText     =   "Frequency"
         Top             =   1380
         Width           =   1155
      End
      Begin VB.Label lblPeak 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   2
         Left            =   3240
         TabIndex        =   10
         ToolTipText     =   "Peak Level (db)"
         Top             =   1020
         Width           =   1155
      End
      Begin VB.Label lblFreq 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   2
         Left            =   480
         TabIndex        =   9
         ToolTipText     =   "Frequency"
         Top             =   1020
         Width           =   1155
      End
      Begin VB.Label lblPeak 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   1
         Left            =   3240
         TabIndex        =   8
         ToolTipText     =   "Peak Level (db)"
         Top             =   660
         Width           =   1155
      End
      Begin VB.Label lblFreq 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   1
         Left            =   480
         TabIndex        =   7
         ToolTipText     =   "Frequency"
         Top             =   660
         Width           =   1155
      End
      Begin VB.Label lblPeak 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   0
         Left            =   3240
         TabIndex        =   6
         ToolTipText     =   "Peak Level (db)"
         Top             =   300
         Width           =   1155
      End
      Begin VB.Label lblFreq 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   0
         Left            =   480
         TabIndex        =   5
         ToolTipText     =   "Frequency"
         Top             =   300
         Width           =   1155
      End
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Presets"
      Height          =   375
      Left            =   6540
      TabIndex        =   108
      Top             =   5160
      Width           =   1395
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Make"
      Height          =   375
      Left            =   1800
      TabIndex        =   50
      Top             =   5160
      Width           =   1395
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Configure"
      Height          =   375
      Left            =   8220
      TabIndex        =   49
      Top             =   5160
      Width           =   1395
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Quit"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   5160
      Width           =   1395
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000C&
      Caption         =   "Wave Form Presets"
      Height          =   4575
      Index           =   3
      Left            =   1800
      TabIndex        =   109
      Top             =   240
      Width           =   5415
      Begin VB.CommandButton Command5 
         Caption         =   "Reset"
         Height          =   375
         Left            =   3840
         TabIndex        =   124
         Top             =   300
         Width           =   1395
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H8000000C&
         Caption         =   "Simple Sine Wave"
         Height          =   255
         Index           =   0
         Left            =   300
         TabIndex        =   123
         Top             =   420
         Value           =   -1  'True
         Width           =   2295
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H8000000C&
         Caption         =   "Detuned Sine Wave Pair"
         Height          =   255
         Index           =   1
         Left            =   300
         TabIndex        =   122
         Top             =   660
         Width           =   3015
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H8000000C&
         Caption         =   "Saw Tooth Wave"
         Height          =   255
         Index           =   3
         Left            =   300
         TabIndex        =   121
         Top             =   1320
         Width           =   2115
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H8000000C&
         Caption         =   "Saw Tooth with Detuned Harmonics"
         Height          =   255
         Index           =   4
         Left            =   300
         TabIndex        =   120
         Top             =   1560
         Width           =   4695
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H8000000C&
         Caption         =   "Saw Tooth with Modulated Harmonics"
         Height          =   255
         Index           =   5
         Left            =   300
         TabIndex        =   119
         Top             =   1800
         Width           =   4695
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H8000000C&
         Caption         =   "Square Wave"
         Height          =   255
         Index           =   6
         Left            =   300
         TabIndex        =   118
         Top             =   2220
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H8000000C&
         Caption         =   "Square with Detuned Harmonics"
         Height          =   255
         Index           =   7
         Left            =   300
         TabIndex        =   117
         Top             =   2460
         Width           =   4695
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H8000000C&
         Caption         =   "Square with Modulated Harmonics"
         Height          =   255
         Index           =   8
         Left            =   300
         TabIndex        =   116
         Top             =   2700
         Width           =   4695
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   240
         TabIndex        =   115
         Top             =   4020
         Width           =   1395
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Accept"
         Height          =   375
         Left            =   3780
         TabIndex        =   114
         Top             =   4020
         Width           =   1395
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H8000000C&
         Caption         =   "Telephone Dial Tone"
         Height          =   255
         Index           =   9
         Left            =   300
         TabIndex        =   113
         Top             =   3120
         Width           =   2715
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H8000000C&
         Caption         =   "Telephone Busy Tone (non-modulated)"
         Height          =   255
         Index           =   10
         Left            =   300
         TabIndex        =   112
         Top             =   3360
         Width           =   4695
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H8000000C&
         Caption         =   "Telephone Ring Tone (non-modulated)"
         Height          =   255
         Index           =   11
         Left            =   300
         TabIndex        =   111
         Top             =   3600
         Width           =   4695
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H8000000C&
         Caption         =   "Detuned Sine Wave Cluster"
         Height          =   255
         Index           =   2
         Left            =   300
         TabIndex        =   110
         Top             =   900
         Width           =   3375
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000C&
      Caption         =   "Configuration"
      Height          =   4095
      Index           =   0
      Left            =   420
      TabIndex        =   1
      Top             =   360
      Width           =   9075
      Begin MSComctlLib.Slider sldRate 
         Height          =   375
         Index           =   0
         Left            =   660
         TabIndex        =   51
         Top             =   1500
         Width           =   3735
         _ExtentX        =   6583
         _ExtentY        =   656
         _Version        =   393216
         Min             =   1
         Max             =   100
         SelStart        =   10
         TickFrequency   =   10
         Value           =   10
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Close"
         Height          =   375
         Left            =   3900
         TabIndex        =   44
         Top             =   3540
         Width           =   1395
      End
      Begin VB.ComboBox boxRate 
         Height          =   360
         Left            =   1380
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   720
         Width           =   2952
      End
      Begin VB.ComboBox boxWidth 
         Height          =   360
         Left            =   4680
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   720
         Width           =   2952
      End
      Begin MSComctlLib.Slider sldDepth 
         Height          =   375
         Index           =   0
         Left            =   4620
         TabIndex        =   52
         Top             =   1500
         Width           =   3735
         _ExtentX        =   6583
         _ExtentY        =   656
         _Version        =   393216
         Min             =   1
         Max             =   100
         SelStart        =   10
         TickFrequency   =   10
         Value           =   10
      End
      Begin MSComctlLib.Slider sldRate 
         Height          =   375
         Index           =   1
         Left            =   660
         TabIndex        =   57
         Top             =   2220
         Width           =   3735
         _ExtentX        =   6583
         _ExtentY        =   656
         _Version        =   393216
         Min             =   1
         Max             =   100
         SelStart        =   10
         TickFrequency   =   10
         Value           =   10
      End
      Begin MSComctlLib.Slider sldDepth 
         Height          =   375
         Index           =   1
         Left            =   4620
         TabIndex        =   58
         Top             =   2220
         Width           =   3735
         _ExtentX        =   6583
         _ExtentY        =   656
         _Version        =   393216
         Min             =   1
         Max             =   100
         SelStart        =   10
         TickFrequency   =   10
         Value           =   10
      End
      Begin MSComctlLib.Slider sldRate 
         Height          =   375
         Index           =   2
         Left            =   660
         TabIndex        =   63
         Top             =   3000
         Width           =   3735
         _ExtentX        =   6583
         _ExtentY        =   656
         _Version        =   393216
         Min             =   1
         Max             =   100
         SelStart        =   10
         TickFrequency   =   10
         Value           =   10
      End
      Begin MSComctlLib.Slider sldDepth 
         Height          =   375
         Index           =   2
         Left            =   4620
         TabIndex        =   64
         Top             =   3000
         Width           =   3735
         _ExtentX        =   6583
         _ExtentY        =   656
         _Version        =   393216
         Min             =   1
         Max             =   100
         SelStart        =   10
         TickFrequency   =   10
         Value           =   10
      End
      Begin VB.Label lblDepth 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "100"
         Height          =   255
         Index           =   2
         Left            =   6120
         TabIndex        =   68
         Top             =   2760
         Width           =   555
      End
      Begin VB.Label lblRate 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "10.0"
         Height          =   255
         Index           =   2
         Left            =   2040
         TabIndex        =   67
         Top             =   2760
         Width           =   615
      End
      Begin VB.Label lblD 
         BackStyle       =   0  'Transparent
         Caption         =   "LFO-3 Depth:"
         Height          =   255
         Index           =   2
         Left            =   4620
         TabIndex        =   66
         Top             =   2760
         Width           =   1455
      End
      Begin VB.Label lblR 
         BackStyle       =   0  'Transparent
         Caption         =   "LFO-3 Rate:"
         Height          =   255
         Index           =   2
         Left            =   660
         TabIndex        =   65
         Top             =   2760
         Width           =   1395
      End
      Begin VB.Label lblDepth 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "100"
         Height          =   255
         Index           =   1
         Left            =   6120
         TabIndex        =   62
         Top             =   1980
         Width           =   555
      End
      Begin VB.Label lblRate 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "10.0"
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   61
         Top             =   1980
         Width           =   615
      End
      Begin VB.Label lblD 
         BackStyle       =   0  'Transparent
         Caption         =   "LFO-2 Depth:"
         Height          =   255
         Index           =   1
         Left            =   4620
         TabIndex        =   60
         Top             =   1980
         Width           =   1455
      End
      Begin VB.Label lblR 
         BackStyle       =   0  'Transparent
         Caption         =   "LFO-2 Rate:"
         Height          =   255
         Index           =   1
         Left            =   660
         TabIndex        =   59
         Top             =   1980
         Width           =   1335
      End
      Begin VB.Label lblDepth 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "100"
         Height          =   255
         Index           =   0
         Left            =   6120
         TabIndex        =   56
         Top             =   1260
         Width           =   555
      End
      Begin VB.Label lblRate 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "10.0"
         Height          =   255
         Index           =   0
         Left            =   2040
         TabIndex        =   55
         Top             =   1260
         Width           =   615
      End
      Begin VB.Label lblD 
         BackStyle       =   0  'Transparent
         Caption         =   "LFO-1 Depth:"
         Height          =   255
         Index           =   0
         Left            =   4620
         TabIndex        =   54
         Top             =   1260
         Width           =   1455
      End
      Begin VB.Label lblR 
         BackStyle       =   0  'Transparent
         Caption         =   "LFO-1 Rate:"
         Height          =   255
         Index           =   0
         Left            =   660
         TabIndex        =   53
         Top             =   1260
         Width           =   1395
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Band Width"
         Height          =   255
         Left            =   4680
         TabIndex        =   46
         Top             =   420
         Width           =   2775
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Sample Rate"
         Height          =   255
         Left            =   1380
         TabIndex        =   45
         Top             =   420
         Width           =   2775
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0000C000&
      Caption         =   "Generator Progress"
      ForeColor       =   &H00C00000&
      Height          =   615
      Index           =   2
      Left            =   1740
      TabIndex        =   47
      Top             =   4020
      Width           =   6615
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   315
         Left            =   0
         TabIndex        =   48
         Top             =   240
         Width           =   1635
      End
   End
End
Attribute VB_Name = "frmPcm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Code module and project created by "Urthman"
'   http://www.jsent.biz/urthman/
'   http://www.mp3.com/urthman/

Option Explicit

Private Const OffColor& = &HC0C0FF
Private Const OnnColor& = &HC0FFC0

Private Const Duration& = 4000

Public Sub CheckRate()

Dim cIndx%
Dim SR&

    SR = Val(Trim(boxRate.List(boxRate.ListIndex)))
    
    For cIndx = 0 To lblFreq.UBound
        If (Val(lblFreq(cIndx).Caption) > (SR / 2)) Then
            chkChan(cIndx).Value = 0
            chkChan(cIndx).Enabled = False
            lblFreq(cIndx).Enabled = False
            chkLFO1(cIndx).Value = 0
            chkLFO1(cIndx).Enabled = False
            chkLFO2(cIndx).Value = 0
            chkLFO2(cIndx).Enabled = False
            chkLFO3(cIndx).Value = 0
            chkLFO3(cIndx).Enabled = False
        Else
            chkChan(cIndx).Enabled = True
            lblFreq(cIndx).Enabled = True
            chkLFO1(cIndx).Enabled = True
            chkLFO1(cIndx).ToolTipText = ("LFO-1 : " & lblRate(0).Caption & "Hz x " & lblDepth(0).Caption)
            chkLFO2(cIndx).Enabled = True
            chkLFO2(cIndx).ToolTipText = ("LFO-2 : " & lblRate(1).Caption & "Hz x " & lblDepth(1).Caption)
            chkLFO3(cIndx).Enabled = True
            chkLFO3(cIndx).ToolTipText = ("LFO-3 : " & lblRate(2).Caption & "Hz x " & lblDepth(2).Caption)
        End If
    Next

End Sub
Public Sub InitForm()

    boxRate.Clear
    boxRate.AddItem "96000"
    boxRate.AddItem "88200"
    boxRate.AddItem "64000"
    boxRate.AddItem "48000"
    boxRate.AddItem "44100"
    boxRate.AddItem "32000"
    boxRate.AddItem "22050"
    boxRate.AddItem "16000"
    boxRate.AddItem "11025"
    boxRate.AddItem " 8000"
    boxRate.AddItem " 6000"
    
    boxWidth.Clear
    boxWidth.AddItem " 8-bit"
    boxWidth.AddItem "16-bit"
    boxWidth.AddItem "24-bit"
    
    lblFreq(0).Caption = 440
    
    ResetWave
    
End Sub
Public Sub ResetWave()

Dim xIndx%, xFreq&
    
Dim FirstFreq&

    FirstFreq = Val(lblFreq(0).Caption)
    
    boxRate.ListIndex = 4
    boxWidth.ListIndex = 1

    Call CheckRate

    For xIndx = 0 To lblFreq.UBound
        xFreq = (FirstFreq * (xIndx + 1))
        lblFreq(xIndx).Caption = xFreq
        chkChan(xIndx).Value = 0
        
        chkLFO1(xIndx).Visible = (chkChan(xIndx).Value = 1)
        chkLFO2(xIndx).Visible = (chkChan(xIndx).Value = 1)
        chkLFO3(xIndx).Visible = (chkChan(xIndx).Value = 1)
    Next

    For xIndx = 0 To lblRate.UBound
        sldRate(xIndx).Value = 1
        sldDepth(xIndx).Value = 1
    Next

    chkChan(0).Value = 1

End Sub
Private Sub SetAdsr(Which%, SetOn As Boolean)

    If SetOn Then
        lblAttack(Which).Caption = 0
        lblPeak(Which).Caption = 0
        lblDecay(Which).Caption = 0
        lblSustain(Which).Caption = 0
        lblRelease(Which).Caption = 0
    Else
        lblAttack(Which).Caption = vbNullString
        lblPeak(Which).Caption = vbNullString
        lblDecay(Which).Caption = vbNullString
        lblSustain(Which).Caption = vbNullString
        lblRelease(Which).Caption = vbNullString
    End If

End Sub
Private Sub SwitchFrame(FromWhat%, ToWhat%)

    Frame1(ToWhat).Enabled = True
    Frame1(ToWhat).ZOrder 0
    Frame1(ToWhat).Refresh
    
    Frame1(FromWhat).Enabled = False

    Command1.Enabled = (ToWhat = 1)
    Command2.Enabled = (ToWhat <> 2)
    Command3.Enabled = (ToWhat = 1)
    Command6.Enabled = (ToWhat = 1)

End Sub
Private Sub chkChan_Click(Index As Integer)

    chkChan(Index).Caption = chkChan(Index).Value
    
    If (chkChan(Index).Value = 1) Then
        chkChan(Index).BackColor = OnnColor
        SetAdsr Index, True
    Else
        chkChan(Index).BackColor = OffColor
        chkLFO1(Index).Value = 0
        chkLFO2(Index).Value = 0
        chkLFO3(Index).Value = 0
        SetAdsr Index, False
    End If

    chkLFO1(Index).Visible = (chkChan(Index).Value = 1)
    chkLFO2(Index).Visible = (chkChan(Index).Value = 1)
    chkLFO3(Index).Visible = (chkChan(Index).Value = 1)

End Sub
Private Sub chkLFO1_Click(Index As Integer)

    If (chkLFO1(Index).Value = 1) And (chkChan(Index).Value = 0) Then
        chkLFO1(Index).Value = 0
        Exit Sub
    End If
    
    chkLFO1(Index).Caption = chkLFO1(Index).Value
    If (chkLFO1(Index).Value = 1) Then
        chkLFO2(Index).Value = 0
        chkLFO3(Index).Value = 0
    End If

End Sub
Private Sub chkLFO2_Click(Index As Integer)

    If (chkLFO2(Index).Value = 1) And (chkChan(Index).Value = 0) Then
        chkLFO2(Index).Value = 0
        Exit Sub
    End If

    chkLFO2(Index).Caption = chkLFO2(Index).Value
    If (chkLFO2(Index).Value = 1) Then
        chkLFO1(Index).Value = 0
        chkLFO3(Index).Value = 0
    End If

End Sub
Private Sub chkLFO3_Click(Index As Integer)

    If (chkLFO3(Index).Value = 1) And (chkChan(Index).Value = 0) Then
        chkLFO3(Index).Value = 0
        Exit Sub
    End If
    
    chkLFO3(Index).Caption = chkLFO3(Index).Value
    If (chkLFO3(Index).Value = 1) Then
        chkLFO1(Index).Value = 0
        chkLFO2(Index).Value = 0
    End If

End Sub
Private Sub Command1_Click()

Dim WaveID%, xIndx%
Dim Freq As Double
Dim BW As BandWidth
Dim SR As Long
Dim MsgRet As VbMsgBoxResult

    Label3.Width = 0
    SwitchFrame 1, 2
    Enabled = False
    
    BW = ((boxWidth.ListIndex + 1) * 8)
    SR = Val(Trim(boxRate.List(boxRate.ListIndex)))
    
    InitRiff SR, BW
    
    For xIndx = 0 To lblFreq.UBound
        Label3.Width = ((xIndx / lblFreq.UBound) * Frame1(2).Width)
        If (lblFreq(xIndx).Caption > (SR / 2)) Then Exit For
        If (chkChan(xIndx).Value = 1) Then
            If (chkLFO1(xIndx).Value = 1) Then
                WaveID = NextSineMod(lblFreq(xIndx).Caption, Duration, lblRate(0).Caption, lblDepth(0).Caption)
            ElseIf (chkLFO2(xIndx).Value = 1) Then
                WaveID = NextSineMod(lblFreq(xIndx).Caption, Duration, lblRate(1).Caption, lblDepth(1).Caption)
            ElseIf (chkLFO3(xIndx).Value = 1) Then
                WaveID = NextSineMod(lblFreq(xIndx).Caption, Duration, lblRate(2).Caption, lblDepth(2).Caption)
            Else
                WaveID = NextSine(lblFreq(xIndx).Caption, Duration, 0)
            End If
            Envelope WaveID, lblAttack(xIndx).Caption, lblPeak(xIndx).Caption, lblDecay(xIndx).Caption, lblSustain(xIndx).Caption, lblRelease(xIndx).Caption
        End If
    Next

    Label3.Width = Frame1(2).Width

    WaveID = (WaveID + 1)
    If (WaveID > 1) Then
        Label3.Caption = "Mixing " & Trim(Str(WaveID)) & " waves"
    Else
        Label3.Caption = "Preparing a single wave"
    End If
    Label3.Refresh
    
    Call MixWaves(AttenuationValue(3))

    Label3.Caption = "Writing audio PCM file"
    Label3.Refresh
    
    LoadRiff OutWave.Sample
    SaveRiff WavFileName

    Label3.Caption = (Trim(Str(bWidth)) & "-bit samples at " & Format(bRate, "#,##0") & " KHz")
    Label3.Refresh

    MsgBox "PCM file generated!" & vbCrLf & vbCrLf & WavFileName, vbInformation, "PCM/Riff Wave"

    Enabled = True
    SwitchFrame 2, 1
    Label3.Caption = vbNullString
    
End Sub
Private Sub Command2_Click()

    End

End Sub
Private Sub Command3_Click()

    SwitchFrame 1, 0
    
End Sub
Private Sub Command4_Click()

    SwitchFrame 0, 1
    Call CheckRate

End Sub
Private Sub Command5_Click()

    Option1(0).Value = True
    InitForm
    SwitchFrame 3, 1

End Sub
Private Sub Command6_Click()

    SwitchFrame 1, 3

End Sub
Private Sub Command7_Click()

    Option1(0).Value = True
    SwitchFrame 3, 1

End Sub
Private Sub Command8_Click()

Dim oIndx%
Dim oFreq%

    For oIndx = 0 To Option1.UBound
        If Option1(oIndx).Value Then Exit For
    Next

    If Not Option1(oIndx).Value Then
        Option1(0).Value = True
        SwitchFrame 3, 1
        Exit Sub
    End If
    
    ResetWave
    
    Select Case oIndx
    Case 1
        lblFreq(1).Caption = (lblFreq(0).Caption + 8)
        chkChan(1).Value = 1
    Case 2
        lblFreq(1).Caption = (lblFreq(0).Caption + 5)
        chkChan(1).Value = 1
        lblFreq(2).Caption = (lblFreq(0).Caption + 10)
        chkChan(2).Value = 1
        lblFreq(3).Caption = (lblFreq(0).Caption + 20)
        chkChan(3).Value = 1
    Case 3
        For oIndx = 1 To lblFreq.UBound
            chkChan(oIndx).Value = 1
            lblFreq(oIndx).Caption = (lblFreq(0).Caption * (oIndx + 1))
            lblAttack(oIndx).Caption = (oIndx * 100)
            lblPeak(oIndx).Caption = (0 - (oIndx * 9))
            lblDecay(oIndx).Caption = (oIndx * 100)
            lblSustain(oIndx).Caption = (0 - (oIndx * 18))
            lblRelease(oIndx).Caption = (oIndx * 100)
        Next
        lblAttack(0).Caption = lblAttack(6).Caption
        lblRelease(0).Caption = lblRelease(6).Caption
    Case 4
        For oIndx = 1 To lblFreq.UBound
            chkChan(oIndx).Value = 1
            lblFreq(oIndx).Caption = ((lblFreq(0).Caption * (oIndx + 1)) - (oIndx * 12))
            lblAttack(oIndx).Caption = (oIndx * 100)
            lblPeak(oIndx).Caption = (0 - (oIndx * 9))
            lblDecay(oIndx).Caption = (oIndx * 100)
            lblSustain(oIndx).Caption = (0 - (oIndx * 18))
            lblRelease(oIndx).Caption = (oIndx * 100)
        Next
        lblAttack(0).Caption = lblAttack(6).Caption
        lblRelease(0).Caption = lblRelease(6).Caption
    Case 5
        For oIndx = 1 To lblFreq.UBound
            chkChan(oIndx).Value = 1
            lblFreq(oIndx).Caption = (lblFreq(0).Caption * (oIndx + 1))
            lblAttack(oIndx).Caption = (oIndx * 100)
            lblPeak(oIndx).Caption = (0 - (oIndx * 9))
            lblDecay(oIndx).Caption = (oIndx * 100)
            lblSustain(oIndx).Caption = (0 - (oIndx * 18))
            lblRelease(oIndx).Caption = (oIndx * 100)
            Select Case oIndx
            Case 1, 4, 7, 10: chkLFO1(oIndx).Value = 1
            Case 2, 5, 8, 11: chkLFO2(oIndx).Value = 1
            Case 3, 6, 9, 12: chkLFO3(oIndx).Value = 1
            End Select
        Next
        sldRate(0).Value = 10
        sldDepth(0).Value = 70
        sldRate(1).Value = 20
        sldDepth(1).Value = 80
        sldRate(2).Value = 30
        sldDepth(2).Value = 90
        lblAttack(0).Caption = lblAttack(6).Caption
        lblRelease(0).Caption = lblRelease(6).Caption
    Case 6
        For oIndx = 1 To lblFreq.UBound
            If ((oIndx / 2) = Int(oIndx / 2)) Then
                chkChan(oIndx).Value = 1
                lblFreq(oIndx).Caption = (lblFreq(0).Caption * (oIndx + 1))
                lblAttack(oIndx).Caption = (oIndx * 100)
                lblPeak(oIndx).Caption = (0 - (oIndx * 9))
                lblDecay(oIndx).Caption = (oIndx * 100)
                lblSustain(oIndx).Caption = (0 - (oIndx * 18))
                lblRelease(oIndx).Caption = (oIndx * 100)
            End If
        Next
        lblAttack(0).Caption = lblAttack(6).Caption
        lblRelease(0).Caption = lblRelease(6).Caption
    Case 7
        For oIndx = 1 To lblFreq.UBound
            If ((oIndx / 2) = Int(oIndx / 2)) Then
                chkChan(oIndx).Value = 1
                lblFreq(oIndx).Caption = ((lblFreq(0).Caption * (oIndx + 1)) - (oIndx * 12))
                lblAttack(oIndx).Caption = (oIndx * 100)
                lblPeak(oIndx).Caption = (0 - (oIndx * 9))
                lblDecay(oIndx).Caption = (oIndx * 100)
                lblSustain(oIndx).Caption = (0 - (oIndx * 18))
                lblRelease(oIndx).Caption = (oIndx * 100)
            End If
        Next
        lblAttack(0).Caption = lblAttack(6).Caption
        lblRelease(0).Caption = lblRelease(6).Caption
    Case 8
        For oIndx = 1 To lblFreq.UBound
            If ((oIndx / 2) = Int(oIndx / 2)) Then
                chkChan(oIndx).Value = 1
                lblFreq(oIndx).Caption = (lblFreq(0).Caption * (oIndx + 1))
                lblAttack(oIndx).Caption = (oIndx * 100)
                lblPeak(oIndx).Caption = (0 - (oIndx * 9))
                lblDecay(oIndx).Caption = (oIndx * 100)
                lblSustain(oIndx).Caption = (0 - (oIndx * 18))
                lblRelease(oIndx).Caption = (oIndx * 100)
                Select Case oIndx
                Case 1, 4, 7, 10: chkLFO1(oIndx).Value = 1
                Case 2, 5, 8, 11: chkLFO2(oIndx).Value = 1
                Case 3, 6, 9, 12: chkLFO3(oIndx).Value = 1
                End Select
            End If
        Next
        sldRate(0).Value = 10
        sldDepth(0).Value = 70
        sldRate(1).Value = 20
        sldDepth(1).Value = 80
        sldRate(2).Value = 30
        sldDepth(2).Value = 90
        lblAttack(0).Caption = lblAttack(6).Caption
        lblRelease(0).Caption = lblRelease(6).Caption
    Case 9
        lblFreq(0).Caption = DT1
        lblFreq(1).Caption = DT2
        chkChan(1).Value = 1
    Case 10
        lblFreq(0).Caption = BT1
        lblFreq(1).Caption = BT2
        chkChan(1).Value = 1
    Case 11
        lblFreq(0).Caption = RT1
        lblFreq(1).Caption = RT2
        chkChan(1).Value = 1
    End Select

    SwitchFrame 3, 1

End Sub
Private Sub Form_Load()

Dim cIndx%

    Top = (Screen.Height - Height) / 2
    Left = (Screen.Width - Width) / 2

    Frame1(0).Top = (Frame1(1).Height - Frame1(0).Height) / 2
    Frame1(0).Left = (Width - Frame1(0).Width) / 2

    Frame1(3).Top = (Frame1(1).Height - Frame1(3).Height) / 2
    Frame1(3).Left = (Width - Frame1(3).Width) / 2

    Frame1(1).Enabled = True
    Frame1(1).ZOrder 0
    Frame1(1).Refresh
    
    Frame1(0).Enabled = False
    Frame1(2).Enabled = False

    InitForm

End Sub
Private Sub lblAttack_DblClick(Index As Integer)

    If (chkChan(Index).Value = 0) Then chkChan(Index).Value = 1
    DoEvents
    Enabled = False
    Load frmAdsr
    frmAdsr.SetMeUp Index
    frmAdsr.Show
    frmAdsr.SetFocus

End Sub


Private Sub lblDecay_DblClick(Index As Integer)

    If (chkChan(Index).Value = 0) Then chkChan(Index).Value = 1
    DoEvents
    Enabled = False
    Load frmAdsr
    frmAdsr.SetMeUp Index
    frmAdsr.Show
    frmAdsr.SetFocus

End Sub


Private Sub lblFreq_DblClick(Index As Integer)

    Load frmHertz
    Call frmHertz.PrepMe(Index)
    frmHertz.Show
    frmHertz.SetFocus
    Enabled = False

End Sub

Private Sub lblPeak_DblClick(Index As Integer)

    If (chkChan(Index).Value = 0) Then chkChan(Index).Value = 1
    DoEvents
    Enabled = False
    Load frmAdsr
    frmAdsr.SetMeUp Index
    frmAdsr.Show
    frmAdsr.SetFocus

End Sub

Private Sub lblRelease_DblClick(Index As Integer)

    If (chkChan(Index).Value = 0) Then chkChan(Index).Value = 1
    DoEvents
    Enabled = False
    Load frmAdsr
    frmAdsr.SetMeUp Index
    frmAdsr.Show
    frmAdsr.SetFocus

End Sub


Private Sub lblSustain_DblClick(Index As Integer)

    If (chkChan(Index).Value = 0) Then chkChan(Index).Value = 1
    DoEvents
    Enabled = False
    Load frmAdsr
    frmAdsr.SetMeUp Index
    frmAdsr.Show
    frmAdsr.SetFocus

End Sub


Private Sub sldDepth_Change(Index As Integer)

    lblDepth(Index).Caption = sldDepth(Index).Value

End Sub
Private Sub sldDepth_Click(Index As Integer)

    lblDepth(Index).Caption = sldDepth(Index).Value

End Sub
Private Sub sldRate_Change(Index As Integer)

    lblRate(Index).Caption = sldRate(Index).Value / 20

End Sub
Private Sub sldRate_Click(Index As Integer)

    lblRate(Index).Caption = sldRate(Index).Value / 20

End Sub
