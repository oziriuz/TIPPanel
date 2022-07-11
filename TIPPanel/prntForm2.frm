VERSION 5.00
Begin VB.Form prntForm2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Форма 2"
   ClientHeight    =   16080
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11700
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   16080
   ScaleWidth      =   11700
   StartUpPosition =   2  'CenterScreen
   Begin VB.VScrollBar VScroll 
      Height          =   11760
      LargeChange     =   500
      Left            =   11280
      Max             =   16000
      SmallChange     =   100
      TabIndex        =   0
      Top             =   0
      Width           =   375
   End
   Begin VB.Frame frmForm2 
      BorderStyle     =   0  'None
      Height          =   16000
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11535
      Begin VB.CommandButton btnPrint 
         Caption         =   "OK"
         Height          =   375
         Left            =   9960
         TabIndex        =   178
         Top             =   15480
         Width           =   1475
      End
      Begin VB.TextBox txtClassP 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   4440
         TabIndex        =   118
         TabStop         =   0   'False
         Top             =   3840
         Width           =   735
      End
      Begin VB.TextBox txtChem1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   2640
         TabIndex        =   117
         TabStop         =   0   'False
         Top             =   3480
         Width           =   2175
      End
      Begin VB.TextBox txtMixTime 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   7560
         TabIndex        =   116
         TabStop         =   0   'False
         Top             =   3840
         Width           =   615
      End
      Begin VB.TextBox txtChem2 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   5760
         TabIndex        =   115
         TabStop         =   0   'False
         Top             =   3480
         Width           =   2175
      End
      Begin VB.TextBox txtCem2 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   5760
         TabIndex        =   114
         TabStop         =   0   'False
         Top             =   3120
         Width           =   2175
      End
      Begin VB.TextBox txtExpTime 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   10560
         TabIndex        =   113
         TabStop         =   0   'False
         Top             =   3840
         Width           =   615
      End
      Begin VB.TextBox txtEDM 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   1800
         TabIndex        =   112
         TabStop         =   0   'False
         Top             =   3840
         Width           =   495
      End
      Begin VB.TextBox txtChem3 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   9000
         TabIndex        =   111
         TabStop         =   0   'False
         Top             =   3480
         Width           =   2175
      End
      Begin VB.TextBox txtCem3 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   9000
         TabIndex        =   110
         TabStop         =   0   'False
         Top             =   3120
         Width           =   2175
      End
      Begin VB.TextBox txtCem1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   2640
         TabIndex        =   109
         TabStop         =   0   'False
         Top             =   3120
         Width           =   2175
      End
      Begin VB.TextBox txtClassH 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   10320
         TabIndex        =   108
         TabStop         =   0   'False
         Top             =   2760
         Width           =   855
      End
      Begin VB.TextBox txtClass 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   1800
         TabIndex        =   107
         TabStop         =   0   'False
         Top             =   2760
         Width           =   855
      End
      Begin VB.TextBox txtOrdVol 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   10080
         TabIndex        =   106
         TabStop         =   0   'False
         Top             =   2400
         Width           =   615
      End
      Begin VB.TextBox txtRecType 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   1680
         TabIndex        =   105
         TabStop         =   0   'False
         Top             =   2400
         Width           =   1695
      End
      Begin VB.TextBox txtDist 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   10080
         TabIndex        =   104
         TabStop         =   0   'False
         Top             =   2040
         Width           =   615
      End
      Begin VB.TextBox txtClassV 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   6720
         TabIndex        =   103
         TabStop         =   0   'False
         Top             =   2760
         Width           =   855
      End
      Begin VB.TextBox txtClassK 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   4320
         TabIndex        =   102
         TabStop         =   0   'False
         Top             =   2760
         Width           =   735
      End
      Begin VB.TextBox txtVol 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   4320
         TabIndex        =   101
         TabStop         =   0   'False
         Top             =   2400
         Width           =   735
      End
      Begin VB.TextBox txtW 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   6720
         TabIndex        =   100
         TabStop         =   0   'False
         Top             =   2400
         Width           =   975
      End
      Begin VB.TextBox txtDrvNo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   6720
         TabIndex        =   99
         TabStop         =   0   'False
         Top             =   2040
         Width           =   1695
      End
      Begin VB.TextBox txtDrv 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   960
         TabIndex        =   98
         TabStop         =   0   'False
         Top             =   2040
         Width           =   4695
      End
      Begin VB.TextBox txtObj 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   6840
         TabIndex        =   97
         TabStop         =   0   'False
         Top             =   1680
         Width           =   4335
      End
      Begin VB.TextBox txtClnt 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   1080
         TabIndex        =   96
         TabStop         =   0   'False
         Top             =   1680
         Width           =   4575
      End
      Begin VB.TextBox txtDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   7440
         TabIndex        =   95
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox txtOrd 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   9240
         TabIndex        =   94
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox txtOper 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   1320
         TabIndex        =   93
         TabStop         =   0   'False
         Top             =   5280
         Width           =   2655
      End
      Begin VB.TextBox txtIMname 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   92
         TabStop         =   0   'False
         Top             =   7200
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.TextBox txtIMname 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   1
         Left            =   360
         TabIndex        =   91
         TabStop         =   0   'False
         Top             =   7560
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.TextBox txtIMname 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   2
         Left            =   360
         TabIndex        =   90
         TabStop         =   0   'False
         Top             =   7920
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.TextBox txtIMname 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   3
         Left            =   360
         TabIndex        =   89
         TabStop         =   0   'False
         Top             =   8280
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.TextBox txtIMname 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   4
         Left            =   360
         TabIndex        =   88
         TabStop         =   0   'False
         Top             =   8640
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.TextBox txtIMkg 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   0
         Left            =   3480
         TabIndex        =   87
         TabStop         =   0   'False
         Top             =   7200
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox txtIMkg 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   1
         Left            =   3480
         TabIndex        =   86
         TabStop         =   0   'False
         Top             =   7560
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox txtIMkgR 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   0
         Left            =   5520
         TabIndex        =   85
         TabStop         =   0   'False
         Top             =   7200
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox txtIMDiff 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   0
         Left            =   7440
         TabIndex        =   84
         TabStop         =   0   'False
         Top             =   7200
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtIMOK 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   0
         Left            =   9240
         TabIndex        =   83
         TabStop         =   0   'False
         Top             =   7200
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtIMkg 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   2
         Left            =   3480
         TabIndex        =   82
         TabStop         =   0   'False
         Top             =   7920
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox txtIMkg 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   3
         Left            =   3480
         TabIndex        =   81
         TabStop         =   0   'False
         Top             =   8280
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox txtIMkg 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   4
         Left            =   3480
         TabIndex        =   80
         TabStop         =   0   'False
         Top             =   8640
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox txtIMkgR 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   1
         Left            =   5520
         TabIndex        =   79
         TabStop         =   0   'False
         Top             =   7560
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox txtIMkgR 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   2
         Left            =   5520
         TabIndex        =   78
         TabStop         =   0   'False
         Top             =   7920
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox txtIMkgR 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   3
         Left            =   5520
         TabIndex        =   77
         TabStop         =   0   'False
         Top             =   8280
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox txtIMkgR 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   4
         Left            =   5520
         TabIndex        =   76
         TabStop         =   0   'False
         Top             =   8640
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox txtIMDiff 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   1
         Left            =   7440
         TabIndex        =   75
         TabStop         =   0   'False
         Top             =   7560
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtIMDiff 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   2
         Left            =   7440
         TabIndex        =   74
         TabStop         =   0   'False
         Top             =   7920
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtIMDiff 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   3
         Left            =   7440
         TabIndex        =   73
         TabStop         =   0   'False
         Top             =   8280
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtIMDiff 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   4
         Left            =   7440
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   8640
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtIMOK 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   1
         Left            =   9240
         TabIndex        =   71
         TabStop         =   0   'False
         Top             =   7560
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtIMOK 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   2
         Left            =   9240
         TabIndex        =   70
         TabStop         =   0   'False
         Top             =   7920
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtIMOK 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   3
         Left            =   9240
         TabIndex        =   69
         TabStop         =   0   'False
         Top             =   8280
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtIMOK 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   4
         Left            =   9240
         TabIndex        =   68
         TabStop         =   0   'False
         Top             =   8640
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtCemname 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   9840
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.TextBox txtCemname 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   1
         Left            =   360
         TabIndex        =   66
         TabStop         =   0   'False
         Top             =   10200
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.TextBox txtCemname 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   2
         Left            =   360
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   10560
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.TextBox txtCemname 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   3
         Left            =   360
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   10920
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.TextBox txtCemkg 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   0
         Left            =   3480
         TabIndex        =   63
         TabStop         =   0   'False
         Top             =   9840
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox txtCemkgR 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   0
         Left            =   5520
         TabIndex        =   62
         TabStop         =   0   'False
         Top             =   9840
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox txtCemDiff 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   0
         Left            =   7440
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   9840
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtCemOK 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   0
         Left            =   9240
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   9840
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtCemkg 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   1
         Left            =   3480
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   10200
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox txtCemkg 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   2
         Left            =   3480
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   10560
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox txtCemkg 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   3
         Left            =   3480
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   10920
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox txtCemkgR 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   1
         Left            =   5520
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   10200
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox txtCemkgR 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   2
         Left            =   5520
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   10560
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox txtCemkgR 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   3
         Left            =   5520
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   10920
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox txtCemDiff 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   1
         Left            =   7440
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   10200
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtCemDiff 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   2
         Left            =   7440
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   10560
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtCemDiff 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   3
         Left            =   7440
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   10920
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtCemOK 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   1
         Left            =   9240
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   10200
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtCemOK 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   2
         Left            =   9240
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   10560
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtCemOK 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   3
         Left            =   9240
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   10920
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtWatkg 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   0
         Left            =   3480
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   11760
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox txtWatkgR 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   0
         Left            =   5520
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   11760
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox txtWatDiff 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   0
         Left            =   7440
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   11760
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtWatOK 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   0
         Left            =   9240
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   11760
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtChemname 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   12960
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.TextBox txtChemname 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   1
         Left            =   360
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   13320
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.TextBox txtChemname 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   2
         Left            =   360
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   13680
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.TextBox txtChemname 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   3
         Left            =   360
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   14040
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.TextBox txtChemname 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   4
         Left            =   360
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   14400
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.TextBox txtChemname 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   5
         Left            =   360
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   14760
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.TextBox txtChemkg 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   0
         Left            =   3480
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   12960
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox txtChemkgR 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   0
         Left            =   5520
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   12960
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox txtChemkg 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   1
         Left            =   3480
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   13320
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox txtChemkg 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   2
         Left            =   3480
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   13680
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox txtChemkg 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   3
         Left            =   3480
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   14040
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox txtChemkg 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   4
         Left            =   3480
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   14400
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox txtChemkgR 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   1
         Left            =   5520
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   13320
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox txtChemkgR 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   2
         Left            =   5520
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   13680
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox txtChemkgR 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   3
         Left            =   5520
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   14040
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox txtChemkgR 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   4
         Left            =   5520
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   14400
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox txtChemkgR 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   5
         Left            =   5520
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   14760
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox txtChemkg 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   5
         Left            =   3480
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   14760
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox txtChemDiff 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   0
         Left            =   7440
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   12960
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtChemOK 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   0
         Left            =   9240
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   12960
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtChemDiff 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   1
         Left            =   7440
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   13320
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtChemDiff 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   2
         Left            =   7440
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   13680
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtChemDiff 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   3
         Left            =   7440
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   14040
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtChemDiff 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   4
         Left            =   7440
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   14400
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtChemDiff 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   5
         Left            =   7440
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   14760
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtChemOK 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   1
         Left            =   9240
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   13320
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtChemOK 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   2
         Left            =   9240
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   13680
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtChemOK 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   3
         Left            =   9240
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   14040
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtChemOK 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   4
         Left            =   9240
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   14400
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtChemOK 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   5
         Left            =   9240
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   14760
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtExpNote 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   19.5
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   480
         Left            =   4440
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   885
         Width           =   2655
      End
      Begin VB.TextBox txtIMname 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   5
         Left            =   360
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   9000
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.TextBox txtIMkg 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   5
         Left            =   3480
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   9000
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox txtIMkgR 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   5
         Left            =   5520
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   9000
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox txtIMDiff 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   5
         Left            =   7440
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   9000
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtIMOK 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   5
         Left            =   9240
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   9000
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtWatkg 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   1
         Left            =   3480
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   12120
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox txtWatkgR 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   1
         Left            =   5520
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   12120
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox txtWatDiff 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   1
         Left            =   7440
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   12120
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtWatOK 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   1
         Left            =   9240
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   12120
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtWatname 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   11760
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.TextBox txtWatname 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   1
         Left            =   360
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   12120
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Label lblConcP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "lblConcP"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   3480
         TabIndex        =   177
         Top             =   0
         Width           =   810
      End
      Begin VB.Label lblFax 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "lblFax"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   6000
         TabIndex        =   176
         Top             =   240
         Width           =   540
      End
      Begin VB.Label lblTel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "lblTel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   6000
         TabIndex        =   175
         Top             =   0
         Width           =   510
      End
      Begin VB.Label lblTown 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "lblTown"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   360
         TabIndex        =   174
         Top             =   240
         Width           =   705
      End
      Begin VB.Label lblCompany 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "lblCompany"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   360
         TabIndex        =   173
         Top             =   0
         Width           =   1080
      End
      Begin VB.Label lblClassP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Водоплътност:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   2880
         TabIndex        =   172
         Top             =   3840
         Width           =   1470
      End
      Begin VB.Label lblCem2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2 - "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   5400
         TabIndex        =   171
         Top             =   3120
         Width           =   255
      End
      Begin VB.Label lblCem1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1 - "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   2280
         TabIndex        =   170
         Top             =   3120
         Width           =   255
      End
      Begin VB.Label lblMM 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "mm"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   2400
         TabIndex        =   169
         Top             =   3840
         Width           =   330
      End
      Begin VB.Label lblChem3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "3 - "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   8640
         TabIndex        =   168
         Top             =   3480
         Width           =   255
      End
      Begin VB.Label lblChem2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2 - "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   5400
         TabIndex        =   167
         Top             =   3480
         Width           =   255
      End
      Begin VB.Label lblChem1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1 - "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   2280
         TabIndex        =   166
         Top             =   3480
         Width           =   255
      End
      Begin VB.Label lblCem3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "3 - "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   8640
         TabIndex        =   165
         Top             =   3120
         Width           =   255
      End
      Begin VB.Label lblM3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "m3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   1
         Left            =   10800
         TabIndex        =   164
         Top             =   2400
         Width           =   270
      End
      Begin VB.Label lblKG 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "kg"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   7680
         TabIndex        =   163
         Top             =   2400
         Width           =   225
      End
      Begin VB.Label lblM3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "m3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   0
         Left            =   5160
         TabIndex        =   162
         Top             =   2400
         Width           =   315
      End
      Begin VB.Label lblOrdVol 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Общо по заявката:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   8160
         TabIndex        =   161
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label lblRecType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Вид разтвор:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   120
         TabIndex        =   160
         Top             =   2400
         Width           =   1470
      End
      Begin VB.Label lblKM 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "km"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   10800
         TabIndex        =   159
         Top             =   2040
         Width           =   270
      End
      Begin VB.Label lblDist 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Разст.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   9240
         TabIndex        =   158
         Top             =   2040
         Width           =   675
      End
      Begin VB.Label lblTimeArr 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Час на пристигане: ......................................"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   5880
         TabIndex        =   157
         Top             =   4440
         Width           =   3630
      End
      Begin VB.Label lblChem 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Химически добавки:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   120
         TabIndex        =   156
         Top             =   3480
         Width           =   1980
      End
      Begin VB.Label lblEDM 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "D max на ЕДМ:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   120
         TabIndex        =   155
         Top             =   3840
         Width           =   1605
      End
      Begin VB.Label lblCem 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Клас цимент:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   120
         TabIndex        =   154
         Top             =   3120
         Width           =   1275
      End
      Begin VB.Label lblClassH 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Клас по с-ние на хлориди:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   7680
         TabIndex        =   153
         Top             =   2760
         Width           =   2505
      End
      Begin VB.Label lblClassV 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Клас по възд.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   5160
         TabIndex        =   152
         Top             =   2760
         Width           =   1380
      End
      Begin VB.Label lblClassK 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Клас по конс.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   2880
         TabIndex        =   151
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label lblClass 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Клас по якост:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   120
         TabIndex        =   150
         Top             =   2760
         Width           =   1590
      End
      Begin VB.Label lblPlaceMix 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Местополагане: .............................................................."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   120
         TabIndex        =   149
         Top             =   4440
         Width           =   4425
      End
      Begin VB.Label lblW 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Тегло:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   6000
         TabIndex        =   148
         Top             =   2400
         Width           =   630
      End
      Begin VB.Label lblVol 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Обем:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   3600
         TabIndex        =   147
         Top             =   2400
         Width           =   645
      End
      Begin VB.Label lblExpTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Час на експедиция:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   8520
         TabIndex        =   146
         Top             =   3840
         Width           =   1890
      End
      Begin VB.Label lblMixTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Начало на миксиране:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   5280
         TabIndex        =   145
         Top             =   3840
         Width           =   2190
      End
      Begin VB.Label lblDrvNo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Кола:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   6000
         TabIndex        =   144
         Top             =   2040
         Width           =   525
      End
      Begin VB.Label lblDrv 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Водач:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   120
         TabIndex        =   143
         Top             =   2040
         Width           =   660
      End
      Begin VB.Label lblObj 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Обект:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   6000
         TabIndex        =   142
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label lblClnt 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Клиент:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   141
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label lblOrd 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "по заявка No. / дата"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   9120
         TabIndex        =   140
         Top             =   720
         Width           =   1965
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "/"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   7200
         TabIndex        =   139
         Top             =   960
         Width           =   105
      End
      Begin VB.Label lblExpNote 
         BackStyle       =   0  'Transparent
         Caption         =   "Експедиционна бележка No."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   138
         Top             =   960
         Width           =   4215
      End
      Begin VB.Label Label35 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Диспечер ТИП-Панел v1.2/2014"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   8640
         TabIndex        =   137
         Top             =   0
         Width           =   2625
      End
      Begin VB.Label lblOperSign 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "/..................."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   4080
         TabIndex        =   136
         Top             =   5280
         Width           =   915
      End
      Begin VB.Label lblSign 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Приел: ....................................."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   8280
         TabIndex        =   135
         Top             =   5280
         Width           =   2385
      End
      Begin VB.Label lblDrvSign 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Водач: ........................."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   5880
         TabIndex        =   134
         Top             =   5280
         Width           =   1830
      End
      Begin VB.Label lblOper 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Диспечер:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   120
         TabIndex        =   133
         Top             =   5280
         Width           =   1005
      End
      Begin VB.Label lblPlaceEnd 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Край на полагане: ........................................"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   5880
         TabIndex        =   132
         Top             =   4800
         Width           =   3630
      End
      Begin VB.Label lblPlaceStart 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Начало на полагането: ...................................."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   120
         TabIndex        =   131
         Top             =   4800
         Width           =   3945
      End
      Begin VB.Label lblIMname 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Инертен материал"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   360
         TabIndex        =   130
         Top             =   6840
         Width           =   1860
      End
      Begin VB.Label lblMkg 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Зададено"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   3480
         TabIndex        =   129
         Top             =   6840
         Width           =   975
      End
      Begin VB.Label lblMkgR 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Изпълнено"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   5520
         TabIndex        =   128
         Top             =   6840
         Width           =   1125
      End
      Begin VB.Label lblMDiff 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Разлика в %"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   7440
         TabIndex        =   127
         Top             =   6840
         Width           =   1230
      End
      Begin VB.Label lblMOK 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Одобрено"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   9240
         TabIndex        =   126
         Top             =   6840
         Width           =   990
      End
      Begin VB.Label lblCemname 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Цимент"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   360
         TabIndex        =   125
         Top             =   9480
         Width           =   735
      End
      Begin VB.Label lblWatname 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Вода"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   360
         TabIndex        =   124
         Top             =   11400
         Width           =   495
      End
      Begin VB.Label lblChemname 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Химически добавки"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   360
         TabIndex        =   123
         Top             =   12600
         Width           =   1935
      End
      Begin VB.Label lblAddPlace 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Място: ................................................"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   120
         TabIndex        =   122
         Top             =   6120
         Width           =   2850
      End
      Begin VB.Label lblAddType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Вид: ..................................."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   3960
         TabIndex        =   121
         Top             =   6120
         Width           =   2040
      End
      Begin VB.Label lblAddSign 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Разпоредител: ........................................"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   7440
         TabIndex        =   120
         Top             =   6120
         Width           =   3330
      End
      Begin VB.Label lblAdd 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Допълнително третиране на разтвора (добавяне на вода или друго) на отговорност на Разпоредителя:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   120
         TabIndex        =   119
         Top             =   5760
         Width           =   10260
      End
   End
End
Attribute VB_Name = "prntForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim i               As Integer
    Dim PrevSet         As Boolean
    Dim strSubKey       As String
    Dim Comp            As String
    Dim Town            As String
    Dim ConcP           As String
    Dim Tel             As String
    Dim Fax             As String
    Dim intEmpFile      As Integer

    Me.Height = 10000
    VScroll.Height = Me.Height - 500
    VScroll.Max = 8000
    VScroll.SmallChange = 500
    VScroll.LargeChange = 1000
    
    intEmpFile = FreeFile
    
    For i = 0 To ns1 - 1
        Me.txtIMname(i).Enabled = True
        Me.txtIMname(i).Visible = True
        Me.txtIMkg(i).Enabled = True
        Me.txtIMkg(i).Visible = True
        Me.txtIMkgR(i).Enabled = True
        Me.txtIMkgR(i).Visible = True
        Me.txtIMDiff(i).Enabled = True
        Me.txtIMDiff(i).Visible = True
        Me.txtIMOK(i).Enabled = True
        Me.txtIMOK(i).Visible = True
    Next i
    For i = 0 To ns3 - 1
        Me.txtCemname(i).Enabled = True
        Me.txtCemname(i).Visible = True
        Me.txtCemkg(i).Enabled = True
        Me.txtCemkg(i).Visible = True
        Me.txtCemkgR(i).Enabled = True
        Me.txtCemkgR(i).Visible = True
        Me.txtCemDiff(i).Enabled = True
        Me.txtCemDiff(i).Visible = True
        Me.txtCemOK(i).Enabled = True
        Me.txtCemOK(i).Visible = True
    Next i
    For i = 0 To ns2 - 1
        Me.txtWatname(i).Enabled = True
        Me.txtWatname(i).Visible = True
        Me.txtWatkg(i).Enabled = True
        Me.txtWatkg(i).Visible = True
        Me.txtWatkgR(i).Enabled = True
        Me.txtWatkgR(i).Visible = True
        Me.txtWatDiff(i).Enabled = True
        Me.txtWatDiff(i).Visible = True
        Me.txtWatOK(i).Enabled = True
        Me.txtWatOK(i).Visible = True
    Next i
    For i = 0 To ns4 - 1
        Me.txtChemname(i).Enabled = True
        Me.txtChemname(i).Visible = True
        Me.txtChemkg(i).Enabled = True
        Me.txtChemkg(i).Visible = True
        Me.txtChemkgR(i).Enabled = True
        Me.txtChemkgR(i).Visible = True
        Me.txtChemDiff(i).Enabled = True
        Me.txtChemDiff(i).Visible = True
        Me.txtChemOK(i).Enabled = True
        Me.txtChemOK(i).Visible = True
    Next i
    
    'зареждане на настройките на форма 2 от регистъра
    strSubKey = Trim(PlaceProgSet2)
    PrevSet = CheckRegistryKey(HKEY_CURRENT_USER, strSubKey)
    
    If PrevSet = True Then
        rDist = GetSetting(PlaceProgSettings, PlaceForm2, "Dist", ErrRes)
        rRecType = GetSetting(PlaceProgSettings, PlaceForm2, "RecType", ErrRes)
        rVol = GetSetting(PlaceProgSettings, PlaceForm2, "Vol", ErrRes)
        rW = GetSetting(PlaceProgSettings, PlaceForm2, "W", ErrRes)
        rOrdVol = GetSetting(PlaceProgSettings, PlaceForm2, "OrdVol", ErrRes)
        rClass = GetSetting(PlaceProgSettings, PlaceForm2, "Class", ErrRes)
        rClassK = GetSetting(PlaceProgSettings, PlaceForm2, "ClassK", ErrRes)
        rClassV = GetSetting(PlaceProgSettings, PlaceForm2, "ClassV", ErrRes)
        rClassH = GetSetting(PlaceProgSettings, PlaceForm2, "ClassH", ErrRes)
        rClassP = GetSetting(PlaceProgSettings, PlaceForm2, "ClassP", ErrRes)
        rCem1 = GetSetting(PlaceProgSettings, PlaceForm2, "Cem1", ErrRes)
        rCem2 = GetSetting(PlaceProgSettings, PlaceForm2, "Cem2", ErrRes)
        rCem3 = GetSetting(PlaceProgSettings, PlaceForm2, "Cem3", ErrRes)
        rChem1 = GetSetting(PlaceProgSettings, PlaceForm2, "Chem1", ErrRes)
        rChem2 = GetSetting(PlaceProgSettings, PlaceForm2, "Chem2", ErrRes)
        rChem3 = GetSetting(PlaceProgSettings, PlaceForm2, "Chem3", ErrRes)
        rEDM = GetSetting(PlaceProgSettings, PlaceForm2, "EDM", ErrRes)
        rMixTime = GetSetting(PlaceProgSettings, PlaceForm2, "MixTime", ErrRes)
        rExpTime = GetSetting(PlaceProgSettings, PlaceForm2, "ExpTime", ErrRes)
        rRealVol = GetSetting(PlaceProgSettings, PlaceForm2, "RealVol", ErrRes)
    Else
        rDist = 1
        rRecType = 1
        rVol = 1
        rW = 1
        rOrdVol = 1
        rClass = 1
        rClassK = 1
        rClassV = 1
        rClassH = 1
        rClassP = 1
        rCem1 = 1
        rCem2 = 1
        rCem3 = 1
        rChem1 = 1
        rChem2 = 1
        rChem3 = 1
        rEDM = 1
        rMixTime = 1
        rExpTime = 1
        rRealVol = 1
    End If
        
    If Dir(InfoFile) <> "" Then
        Open InfoFile For Input As intEmpFile
        Input #intEmpFile, Comp, Town, ConcP, Tel, Fax
        Close
        If Len(Comp) Then
            Me.lblCompany.Caption = Comp
        Else
            Me.lblCompany.Caption = ""
        End If
        If Len(Town) Then
            Me.lblTown.Caption = Town
        Else
            Me.lblTown.Caption = ""
        End If
        If Len(ConcP) Then
            Me.lblConcP.Caption = ConcP
        Else
            Me.lblConcP.Caption = ""
        End If
        If Len(Tel) Then
            Me.lblTel.Caption = Tel
        Else
            Me.lblTel.Caption = ""
        End If
        If Len(Fax) Then
            Me.lblFax.Caption = Fax
        Else
            Me.lblFax.Caption = ""
        End If
    Else
        Me.lblCompany.Caption = ""
        Me.lblTown.Caption = ""
        Me.lblConcP.Caption = ""
        Me.lblTel.Caption = ""
        Me.lblFax.Caption = ""
    End If
    If rDist = 0 Then
        Me.lblDist.Visible = False
        Me.txtDist.Enabled = False
        Me.lblKM.Visible = False
    ElseIf rDist = 1 Then
        Me.lblDist.Visible = True
        Me.txtDist.Enabled = True
        Me.lblKM.Visible = True
    End If
    If rRecType = 0 Then
        Me.lblRecType.Visible = False
        Me.txtRecType.Enabled = False
    ElseIf rRecType = 1 Then
        Me.lblRecType.Visible = True
        Me.txtRecType.Enabled = True
    End If
    If rVol = 0 Then
        Me.lblVol.Visible = False
        Me.txtVol.Enabled = False
        Me.lblM3(0).Visible = False
    ElseIf rVol = 1 Then
        Me.lblVol.Visible = True
        Me.txtVol.Enabled = True
        Me.lblM3(0).Visible = True
    End If
    If rW = 0 Then
        Me.lblW.Visible = False
        Me.txtW.Enabled = False
        Me.lblKG.Visible = False
    ElseIf rW = 1 Then
        Me.lblW.Visible = True
        Me.txtW.Enabled = True
        Me.lblKG.Visible = True
    End If
    If rOrdVol = 0 Then
        Me.lblOrdVol.Visible = False
        Me.txtOrdVol.Enabled = False
        Me.lblM3(1).Visible = False
    ElseIf rOrdVol = 1 Then
        Me.lblOrdVol.Visible = True
        Me.txtOrdVol.Enabled = True
        Me.lblM3(1).Visible = True
    End If
    If rClass = 0 Then
        Me.lblClass.Visible = False
        Me.txtClass.Enabled = False
    ElseIf rClass = 1 Then
        Me.lblClass.Visible = True
        Me.txtClass.Enabled = True
    End If
    If rClassK = 0 Then
        Me.lblClassK.Visible = False
        Me.txtClassK.Enabled = False
    ElseIf rClassK = 1 Then
        Me.lblClassK.Visible = True
        Me.txtClassK.Enabled = True
    End If
    If rClassV = 0 Then
        Me.lblClassV.Visible = False
        Me.txtClassV.Enabled = False
    ElseIf rClassV = 1 Then
        Me.lblClassV.Visible = True
        Me.txtClassV.Enabled = True
    End If
    If rClassH = 0 Then
        Me.lblClassH.Visible = False
        Me.txtClassH.Enabled = False
    ElseIf rClassH = 1 Then
        Me.lblClassH.Visible = True
        Me.txtClassH.Enabled = True
    End If
    If rClassP = 0 Then
        Me.lblClassP.Visible = False
        Me.txtClassP.Enabled = False
    ElseIf rClassP = 1 Then
        Me.lblClassP.Visible = True
        Me.txtClassP.Enabled = True
    End If
    If rCem1 = 0 Then
        Me.lblCem1.Visible = False
        Me.txtCem1.Enabled = False
    ElseIf rCem1 = 1 Then
        Me.lblCem1.Visible = True
        Me.txtCem1.Enabled = True
    End If
    If rCem2 = 0 Then
        Me.lblCem2.Visible = False
        Me.txtCem2.Enabled = False
    ElseIf rCem2 = 1 Then
        Me.lblCem2.Visible = True
        Me.txtCem2.Enabled = True
    End If
    If rCem3 = 0 Then
        Me.lblCem3.Visible = False
        Me.txtCem3.Enabled = False
    ElseIf rCem3 = 1 Then
        Me.lblCem3.Visible = True
        Me.txtCem3.Enabled = True
    End If
    If rCem1 = 0 And rCem2 = 0 And rCem3 = 0 Then
        lblCem.Visible = False
    Else
        lblCem.Visible = True
    End If
    If rChem1 = 0 Then
        Me.lblChem1.Visible = False
        Me.txtChem1.Enabled = False
    ElseIf rChem1 = 1 Then
        Me.lblChem1.Visible = True
        Me.txtChem1.Enabled = True
    End If
    If rChem2 = 0 Then
        Me.lblChem2.Visible = False
        Me.txtChem2.Enabled = False
    ElseIf rChem2 = 1 Then
        Me.lblChem2.Visible = True
        Me.txtChem2.Enabled = True
    End If
    If rChem3 = 0 Then
        Me.lblChem3.Visible = False
        Me.txtChem3.Enabled = False
    ElseIf rChem3 = 1 Then
        Me.lblChem3.Visible = True
        Me.txtChem3.Enabled = True
    End If
    If rChem1 = 0 And rChem2 = 0 And rChem3 = 0 Then
        lblChem.Visible = False
    Else
        lblChem.Visible = True
    End If
    If rEDM = 0 Then
        Me.lblEDM.Visible = False
        Me.txtEDM.Enabled = False
        Me.lblMM.Visible = False
    ElseIf rEDM = 1 Then
        Me.lblEDM.Visible = True
        Me.txtEDM.Enabled = True
        Me.lblMM.Visible = True
    End If
    If rMixTime = 0 Then
        Me.lblMixTime.Visible = False
        Me.txtMixTime.Enabled = False
    ElseIf rMixTime = 1 Then
        Me.lblMixTime.Visible = True
        Me.txtMixTime.Enabled = True
    End If
    If rExpTime = 0 Then
        Me.lblExpTime.Visible = False
        Me.txtExpTime.Enabled = False
    ElseIf rExpTime = 1 Then
        Me.lblExpTime.Visible = True
        Me.txtExpTime.Enabled = True
    End If
End Sub

Private Sub VScroll_Change()

    VScroll.Top = 0: VScroll.Left = Me.ScaleWidth - VScroll.Width
    frmForm2.Top = -1 * VScroll
End Sub

Private Sub btnPrint_Click()

    Call PrintThisForm2(prntForm2)
End Sub

