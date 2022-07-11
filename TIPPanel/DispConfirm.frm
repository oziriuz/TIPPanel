VERSION 5.00
Begin VB.Form DispConfirm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DispConfirm"
   ClientHeight    =   8325
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12855
   Icon            =   "DispConfirm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8325
   ScaleWidth      =   12855
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox initChem 
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
      Height          =   375
      Index           =   5
      Left            =   14400
      Locked          =   -1  'True
      TabIndex        =   133
      TabStop         =   0   'False
      Top             =   5880
      Width           =   615
   End
   Begin VB.TextBox initChem 
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
      Height          =   375
      Index           =   4
      Left            =   14400
      Locked          =   -1  'True
      TabIndex        =   132
      TabStop         =   0   'False
      Top             =   5520
      Width           =   615
   End
   Begin VB.TextBox initChem 
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
      Height          =   375
      Index           =   3
      Left            =   14400
      Locked          =   -1  'True
      TabIndex        =   131
      TabStop         =   0   'False
      Top             =   5160
      Width           =   615
   End
   Begin VB.TextBox initChem 
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
      Height          =   375
      Index           =   2
      Left            =   14400
      Locked          =   -1  'True
      TabIndex        =   130
      TabStop         =   0   'False
      Top             =   4800
      Width           =   615
   End
   Begin VB.TextBox initChem 
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
      Height          =   375
      Index           =   1
      Left            =   14400
      Locked          =   -1  'True
      TabIndex        =   129
      TabStop         =   0   'False
      Top             =   4440
      Width           =   615
   End
   Begin VB.TextBox initChem 
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
      Height          =   375
      Index           =   0
      Left            =   14400
      Locked          =   -1  'True
      TabIndex        =   128
      TabStop         =   0   'False
      Top             =   4080
      Width           =   615
   End
   Begin VB.TextBox initWat 
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
      Height          =   375
      Index           =   1
      Left            =   13800
      Locked          =   -1  'True
      TabIndex        =   127
      TabStop         =   0   'False
      Top             =   4440
      Width           =   615
   End
   Begin VB.TextBox initWat 
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
      Height          =   375
      Index           =   0
      Left            =   13800
      Locked          =   -1  'True
      TabIndex        =   126
      TabStop         =   0   'False
      Top             =   4080
      Width           =   615
   End
   Begin VB.TextBox initCem 
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
      Height          =   375
      Index           =   3
      Left            =   14400
      Locked          =   -1  'True
      TabIndex        =   125
      TabStop         =   0   'False
      Top             =   2640
      Width           =   615
   End
   Begin VB.TextBox initCem 
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
      Height          =   375
      Index           =   2
      Left            =   14400
      Locked          =   -1  'True
      TabIndex        =   124
      TabStop         =   0   'False
      Top             =   2280
      Width           =   615
   End
   Begin VB.TextBox initCem 
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
      Height          =   375
      Index           =   1
      Left            =   14400
      Locked          =   -1  'True
      TabIndex        =   123
      TabStop         =   0   'False
      Top             =   1920
      Width           =   615
   End
   Begin VB.TextBox initCem 
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
      Height          =   375
      Index           =   0
      Left            =   14400
      Locked          =   -1  'True
      TabIndex        =   122
      TabStop         =   0   'False
      Top             =   1560
      Width           =   615
   End
   Begin VB.TextBox initIM 
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
      Height          =   375
      Index           =   5
      Left            =   13800
      Locked          =   -1  'True
      TabIndex        =   121
      TabStop         =   0   'False
      Top             =   3360
      Width           =   615
   End
   Begin VB.TextBox initIM 
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
      Height          =   375
      Index           =   4
      Left            =   13800
      Locked          =   -1  'True
      TabIndex        =   120
      TabStop         =   0   'False
      Top             =   3000
      Width           =   615
   End
   Begin VB.TextBox initIM 
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
      Height          =   375
      Index           =   3
      Left            =   13800
      Locked          =   -1  'True
      TabIndex        =   119
      TabStop         =   0   'False
      Top             =   2640
      Width           =   615
   End
   Begin VB.TextBox initIM 
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
      Height          =   375
      Index           =   2
      Left            =   13800
      Locked          =   -1  'True
      TabIndex        =   118
      TabStop         =   0   'False
      Top             =   2280
      Width           =   615
   End
   Begin VB.TextBox initIM 
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
      Height          =   375
      Index           =   1
      Left            =   13800
      Locked          =   -1  'True
      TabIndex        =   117
      TabStop         =   0   'False
      Top             =   1920
      Width           =   615
   End
   Begin VB.TextBox initIM 
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
      Height          =   375
      Index           =   0
      Left            =   13800
      Locked          =   -1  'True
      TabIndex        =   116
      TabStop         =   0   'False
      Top             =   1560
      Width           =   615
   End
   Begin VB.TextBox txtConfDispDate 
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
      Height          =   375
      Left            =   360
      Locked          =   -1  'True
      TabIndex        =   110
      TabStop         =   0   'False
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox txtConfOrdQuant 
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
      Height          =   375
      Left            =   7440
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   240
      Width           =   615
   End
   Begin VB.Frame frConfDrvInfo 
      Caption         =   "frConfDrvInfo"
      Height          =   2775
      Left            =   8040
      TabIndex        =   85
      Top             =   4560
      Width           =   4695
      Begin VB.TextBox txtConfDrvTel 
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
         Height          =   375
         Left            =   1560
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   2160
         Width           =   1815
      End
      Begin VB.TextBox txtConfDrvMod 
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
         Height          =   375
         Left            =   1560
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   1800
         Width           =   2775
      End
      Begin VB.TextBox txtConfDrvCap 
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
         Height          =   375
         Left            =   1560
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox txtConfDrvReg 
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
         Height          =   375
         Left            =   1560
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtConfDrvName 
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
         Height          =   375
         Left            =   1560
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   720
         Width           =   2775
      End
      Begin VB.TextBox txtConfDrv 
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
         Height          =   375
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblM 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "m3"
         Height          =   195
         Index           =   0
         Left            =   2280
         TabIndex        =   103
         Top             =   1560
         Width           =   210
      End
      Begin VB.Label lblConfDrvName 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblConfDrvName"
         Height          =   255
         Left            =   240
         TabIndex        =   91
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label lblConfDrvReg 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblConfDrvReg"
         Height          =   255
         Left            =   240
         TabIndex        =   90
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label lblConfDrvCap 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblConfDrvCap"
         Height          =   255
         Left            =   240
         TabIndex        =   89
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label lblConfDrvMod 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblConfDrvMod"
         Height          =   255
         Left            =   240
         TabIndex        =   88
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label lblConfDrvTel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblConfDrvTel"
         Height          =   255
         Left            =   240
         TabIndex        =   87
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label lblConfDrv 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblConfDrv"
         Height          =   255
         Left            =   240
         TabIndex        =   86
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Frame frConfClntInfo 
      Caption         =   "frConfClntInfo"
      Height          =   3375
      Left            =   8040
      TabIndex        =   76
      Top             =   1080
      Width           =   4695
      Begin VB.TextBox txtConfClntKm 
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
         Height          =   375
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   2880
         Width           =   615
      End
      Begin VB.TextBox txtConfClntObj 
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
         Height          =   375
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   2520
         Width           =   2775
      End
      Begin VB.TextBox txtConfClntTel 
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
         Height          =   375
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   2160
         Width           =   1815
      End
      Begin VB.TextBox txtConfClntAdd 
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
         Height          =   375
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   1800
         Width           =   2775
      End
      Begin VB.TextBox txtConfClntMOL 
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
         Height          =   375
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   1440
         Width           =   2775
      End
      Begin VB.TextBox txtConfClntBG 
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
         Height          =   375
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox txtConfClntName 
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
         Height          =   375
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   720
         Width           =   2775
      End
      Begin VB.TextBox txtConfClnt 
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
         Height          =   375
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblKm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "km"
         Height          =   195
         Left            =   4080
         TabIndex        =   109
         Top             =   3000
         Width           =   210
      End
      Begin VB.Label lblConfClntObj 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblConfClntObj"
         Height          =   255
         Left            =   240
         TabIndex        =   84
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label lblConfClnt 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblConfClnt"
         Height          =   255
         Left            =   240
         TabIndex        =   83
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label lblConfClntName 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblConfClntName"
         Height          =   255
         Left            =   120
         TabIndex        =   82
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label lblConfClntBG 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblConfClntBG"
         Height          =   255
         Left            =   240
         TabIndex        =   81
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label lblConfClntMOL 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblConfClntMOL"
         Height          =   255
         Left            =   240
         TabIndex        =   80
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label lblConfClntAdd 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblConfClntAdd"
         Height          =   255
         Left            =   240
         TabIndex        =   79
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label lblConfClntTel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblConfClntTel"
         Height          =   255
         Left            =   240
         TabIndex        =   78
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label lblConfClntKm 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblConfClntKm"
         Height          =   255
         Left            =   360
         TabIndex        =   77
         Top             =   3000
         Width           =   2895
      End
   End
   Begin VB.TextBox txtConfCoef 
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
      Height          =   375
      Left            =   7440
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   600
      Width           =   735
   End
   Begin VB.TextBox txtConfDispCount 
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
      Height          =   375
      Left            =   3480
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   600
      Width           =   495
   End
   Begin VB.CommandButton btnSendToController 
      Caption         =   "btnSendToController"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8040
      TabIndex        =   0
      Top             =   7440
      Width           =   4695
   End
   Begin VB.TextBox txtConfDispQuant 
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
      Height          =   375
      Left            =   11520
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   240
      Width           =   615
   End
   Begin VB.Frame frConfRecInfo 
      Caption         =   "frConfRecInfo"
      Height          =   7095
      Left            =   120
      TabIndex        =   63
      Top             =   1080
      Width           =   7815
      Begin VB.TextBox txtConfRec2 
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
         Height          =   375
         Index           =   1
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   115
         TabStop         =   0   'False
         Top             =   6360
         Width           =   2295
      End
      Begin VB.TextBox txtConfRec2 
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
         Height          =   375
         Index           =   0
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   114
         TabStop         =   0   'False
         Top             =   6000
         Width           =   2295
      End
      Begin VB.TextBox txtConfRecKg2 
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
         Height          =   375
         Index           =   1
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   113
         TabStop         =   0   'False
         Top             =   6360
         Width           =   615
      End
      Begin VB.TextBox txtConfRecKg1 
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
         Height          =   375
         Index           =   5
         Left            =   6840
         Locked          =   -1  'True
         TabIndex        =   112
         TabStop         =   0   'False
         Top             =   3240
         Width           =   615
      End
      Begin VB.TextBox txtConfRec1 
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
         Height          =   375
         Index           =   5
         Left            =   4440
         Locked          =   -1  'True
         TabIndex        =   111
         TabStop         =   0   'False
         Top             =   3240
         Width           =   2295
      End
      Begin VB.TextBox txtConfRecClassP 
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
         Height          =   375
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   97
         TabStop         =   0   'False
         Top             =   2880
         Width           =   1335
      End
      Begin VB.TextBox txtConfRecType 
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
         Height          =   375
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1080
         Width           =   2535
      End
      Begin VB.TextBox txtConfRecEDM 
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
         Height          =   375
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   3240
         Width           =   1335
      End
      Begin VB.TextBox txtConfRecClassH 
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
         Height          =   375
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   2520
         Width           =   1335
      End
      Begin VB.TextBox txtConfRecClassV 
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
         Height          =   375
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox txtConfRecClassK 
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
         Height          =   375
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1800
         Width           =   1335
      End
      Begin VB.TextBox txtConfRecKg2 
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
         Height          =   375
         Index           =   0
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   6000
         Width           =   615
      End
      Begin VB.TextBox txtConfRec3 
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
         Height          =   375
         Index           =   3
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   5160
         Width           =   2295
      End
      Begin VB.TextBox txtConfRec3 
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
         Height          =   375
         Index           =   2
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   4800
         Width           =   2295
      End
      Begin VB.TextBox txtConfRec3 
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
         Height          =   375
         Index           =   1
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   4440
         Width           =   2295
      End
      Begin VB.TextBox txtConfRec3 
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
         Height          =   375
         Index           =   0
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   4080
         Width           =   2295
      End
      Begin VB.TextBox txtConfRecKg3 
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
         Height          =   375
         Index           =   3
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   5160
         Width           =   615
      End
      Begin VB.TextBox txtConfRecKg3 
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
         Height          =   375
         Index           =   2
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   4800
         Width           =   615
      End
      Begin VB.TextBox txtConfRecKg3 
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
         Height          =   375
         Index           =   1
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   4440
         Width           =   615
      End
      Begin VB.TextBox txtConfRecKg3 
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
         Height          =   375
         Index           =   0
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   4080
         Width           =   615
      End
      Begin VB.TextBox txtConfRecKg4 
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
         Height          =   375
         Index           =   5
         Left            =   6840
         Locked          =   -1  'True
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   5880
         Width           =   615
      End
      Begin VB.TextBox txtConfRecKg4 
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
         Height          =   375
         Index           =   4
         Left            =   6840
         Locked          =   -1  'True
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   5520
         Width           =   615
      End
      Begin VB.TextBox txtConfRecKg4 
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
         Height          =   375
         Index           =   3
         Left            =   6840
         Locked          =   -1  'True
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   5160
         Width           =   615
      End
      Begin VB.TextBox txtConfRecKg4 
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
         Height          =   375
         Index           =   2
         Left            =   6840
         Locked          =   -1  'True
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   4800
         Width           =   615
      End
      Begin VB.TextBox txtConfRecKg4 
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
         Height          =   375
         Index           =   1
         Left            =   6840
         Locked          =   -1  'True
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   4440
         Width           =   615
      End
      Begin VB.TextBox txtConfRecKg4 
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
         Height          =   375
         Index           =   0
         Left            =   6840
         Locked          =   -1  'True
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   4080
         Width           =   615
      End
      Begin VB.TextBox txtConfRec4 
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
         Height          =   375
         Index           =   5
         Left            =   4440
         Locked          =   -1  'True
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   5880
         Width           =   2295
      End
      Begin VB.TextBox txtConfRec4 
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
         Height          =   375
         Index           =   4
         Left            =   4440
         Locked          =   -1  'True
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   5520
         Width           =   2295
      End
      Begin VB.TextBox txtConfRec4 
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
         Height          =   375
         Index           =   3
         Left            =   4440
         Locked          =   -1  'True
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   5160
         Width           =   2295
      End
      Begin VB.TextBox txtConfRec4 
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
         Height          =   375
         Index           =   2
         Left            =   4440
         Locked          =   -1  'True
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   4800
         Width           =   2295
      End
      Begin VB.TextBox txtConfRec4 
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
         Height          =   375
         Index           =   1
         Left            =   4440
         Locked          =   -1  'True
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   4440
         Width           =   2295
      End
      Begin VB.TextBox txtConfRec4 
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
         Height          =   375
         Index           =   0
         Left            =   4440
         Locked          =   -1  'True
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   4080
         Width           =   2295
      End
      Begin VB.TextBox txtConfRecKg1 
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
         Height          =   375
         Index           =   4
         Left            =   6840
         Locked          =   -1  'True
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   2880
         Width           =   615
      End
      Begin VB.TextBox txtConfRecKg1 
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
         Height          =   375
         Index           =   3
         Left            =   6840
         Locked          =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   2520
         Width           =   615
      End
      Begin VB.TextBox txtConfRecKg1 
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
         Height          =   375
         Index           =   2
         Left            =   6840
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   2160
         Width           =   615
      End
      Begin VB.TextBox txtConfRecKg1 
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
         Height          =   375
         Index           =   1
         Left            =   6840
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   1800
         Width           =   615
      End
      Begin VB.TextBox txtConfRec1 
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
         Height          =   375
         Index           =   4
         Left            =   4440
         Locked          =   -1  'True
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   2880
         Width           =   2295
      End
      Begin VB.TextBox txtConfRec1 
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
         Height          =   375
         Index           =   3
         Left            =   4440
         Locked          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   2520
         Width           =   2295
      End
      Begin VB.TextBox txtConfRec1 
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
         Height          =   375
         Index           =   2
         Left            =   4440
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   2160
         Width           =   2295
      End
      Begin VB.TextBox txtConfRec1 
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
         Height          =   375
         Index           =   1
         Left            =   4440
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   1800
         Width           =   2295
      End
      Begin VB.TextBox txtConfRecKg1 
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
         Height          =   375
         Index           =   0
         Left            =   6840
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox txtConfRec1 
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
         Height          =   375
         Index           =   0
         Left            =   4440
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   1440
         Width           =   2295
      End
      Begin VB.TextBox txtConfRecTimeMix 
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
         Height          =   375
         Left            =   6840
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox txtConfRecTimePour 
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
         Height          =   375
         Left            =   6840
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox txtConfRecClass 
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
         Height          =   375
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox txtConfRecName 
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
         Height          =   375
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   720
         Width           =   2535
      End
      Begin VB.TextBox txtConfRec 
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
         Height          =   375
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblS 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "s"
         Height          =   195
         Index           =   1
         Left            =   7440
         TabIndex        =   108
         Top             =   840
         Width           =   75
      End
      Begin VB.Label lblS 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "s"
         Height          =   195
         Index           =   0
         Left            =   7440
         TabIndex        =   107
         Top             =   480
         Width           =   75
      End
      Begin VB.Label lblKg 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "kg"
         Height          =   195
         Index           =   3
         Left            =   7080
         TabIndex        =   102
         Top             =   3840
         Width           =   180
      End
      Begin VB.Label lblKg 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "kg"
         Height          =   195
         Index           =   2
         Left            =   3600
         TabIndex        =   101
         Top             =   5760
         Width           =   180
      End
      Begin VB.Label lblKg 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "kg"
         Height          =   195
         Index           =   1
         Left            =   3600
         TabIndex        =   100
         Top             =   3840
         Width           =   180
      End
      Begin VB.Label lblKg 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "kg"
         Height          =   195
         Index           =   0
         Left            =   7080
         TabIndex        =   99
         Top             =   1200
         Width           =   180
      End
      Begin VB.Label lblConfRecClassP 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblConfRecClassP"
         Height          =   255
         Left            =   480
         TabIndex        =   98
         Top             =   3000
         Width           =   2295
      End
      Begin VB.Label lblConfRecType 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblConfRecType"
         Height          =   255
         Left            =   120
         TabIndex        =   96
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label lblConfRecClassH 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblConfRecClassH"
         Height          =   255
         Left            =   480
         TabIndex        =   95
         Top             =   2640
         Width           =   2295
      End
      Begin VB.Label lblConfRecClassV 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblConfRecClassV"
         Height          =   255
         Left            =   480
         TabIndex        =   94
         Top             =   2280
         Width           =   2295
      End
      Begin VB.Label lblConfRecClassK 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblConfRecClassK"
         Height          =   255
         Left            =   480
         TabIndex        =   93
         Top             =   1920
         Width           =   2295
      End
      Begin VB.Label lblConfRecChem 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "lblConfRecChem"
         Height          =   255
         Left            =   4440
         TabIndex        =   73
         Top             =   3840
         Width           =   2295
      End
      Begin VB.Label lblConfRecWat 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "lblConfRecWat"
         Height          =   255
         Left            =   960
         TabIndex        =   72
         Top             =   5760
         Width           =   2295
      End
      Begin VB.Label lblConfRecCem 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "lblConfRecCem"
         Height          =   255
         Left            =   960
         TabIndex        =   71
         Top             =   3840
         Width           =   2295
      End
      Begin VB.Label lblConfRecIM 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "lblConfRecIM"
         Height          =   255
         Left            =   4440
         TabIndex        =   70
         Top             =   1200
         Width           =   2295
      End
      Begin VB.Label lblConfRecTimePour 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblConfRecTimePour"
         Height          =   255
         Left            =   5280
         TabIndex        =   69
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label lblConfRecTimeMix 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblConfRecTimeMix"
         Height          =   255
         Left            =   5280
         TabIndex        =   68
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label lblConfRecEDM 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblConfRecEDM"
         Height          =   255
         Left            =   600
         TabIndex        =   67
         Top             =   3360
         Width           =   2175
      End
      Begin VB.Label lblConfRecClass 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblConfRecClass"
         Height          =   255
         Left            =   480
         TabIndex        =   66
         Top             =   1560
         Width           =   2295
      End
      Begin VB.Label lblConfRecName 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblConfRecName"
         Height          =   255
         Left            =   120
         TabIndex        =   65
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label lblConfRec 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblConfRec"
         Height          =   255
         Left            =   120
         TabIndex        =   64
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.TextBox txtConfDisp 
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
      Height          =   375
      Left            =   3480
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   240
      Width           =   975
   End
   Begin VB.Label lblM 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "m3"
      Height          =   195
      Index           =   3
      Left            =   8400
      TabIndex        =   106
      Top             =   720
      Width           =   210
   End
   Begin VB.Label lblM 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "m3"
      Height          =   195
      Index           =   2
      Left            =   12240
      TabIndex        =   105
      Top             =   360
      Width           =   210
   End
   Begin VB.Label lblM 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "m3"
      Height          =   195
      Index           =   1
      Left            =   8160
      TabIndex        =   104
      Top             =   360
      Width           =   210
   End
   Begin VB.Label lblConfOrdQuant 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "lblConfOrdQuant"
      Height          =   255
      Left            =   5520
      TabIndex        =   92
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label lblConfCoef 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "lblConfCoef"
      Height          =   255
      Left            =   4560
      TabIndex        =   75
      Top             =   720
      Width           =   2655
   End
   Begin VB.Label lblConfCount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "lblConfCount"
      Height          =   255
      Left            =   600
      TabIndex        =   74
      Top             =   720
      Width           =   2775
   End
   Begin VB.Label lblConfDispQuant 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "lblConfDispQuant"
      Height          =   255
      Left            =   9000
      TabIndex        =   62
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label lblConfDisp 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "lblConfDisp"
      Height          =   255
      Left            =   2280
      TabIndex        =   61
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "DispConfirm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
'форма за изпращане на данните към контролера

    Dim ctl As Control

    DispConfirm.Caption = frmConfSend
    lblConfDisp.Caption = uniCode
    lblConfOrdQuant.Caption = uniOrdQuant
    lblConfDispQuant.Caption = uniDispQuant
    lblConfCount.Caption = uniNumMix
    lblConfCoef.Caption = uniQuantMix
    frConfRecInfo.Caption = uniInfoMix
    frConfClntInfo.Caption = uniInfoClnt
    frConfDrvInfo.Caption = uniInfoDrv
    lblConfRec.Caption = uniCode
    lblConfRecName.Caption = uniNm
    lblConfRecType.Caption = uniRecType
    lblConfRecClass.Caption = uniClass
    lblConfRecClassK.Caption = uniClassK
    lblConfRecClassV.Caption = uniClassV
    lblConfRecClassH.Caption = uniClassH
    lblConfRecClassP.Caption = uniClassP
    lblConfRecEDM.Caption = uniEDM
    lblConfRecTimePour.Caption = uniTimePour
    lblConfRecTimeMix.Caption = uniTimeMix
    lblConfRecIM.Caption = uniIM
    lblConfRecCem.Caption = uniCem
    lblConfRecWat.Caption = uniWat
    lblConfRecChem.Caption = uniChem
    lblConfClnt.Caption = uniCode
    lblConfClntName.Caption = uniFirm
    lblConfClntBG.Caption = uniBG
    lblConfClntMOL.Caption = uniMOL
    lblConfClntAdd.Caption = uniAdd
    lblConfClntTel.Caption = uniTel
    lblConfClntObj.Caption = uniObj
    lblConfClntKm.Caption = uniKm
    lblConfDrv.Caption = uniCode
    lblConfDrvName.Caption = uniNm
    lblConfDrvReg.Caption = uniDrvReg
    lblConfDrvCap.Caption = uniCapacity
    lblConfDrvMod.Caption = uniMod
    lblConfDrvTel.Caption = uniTel
    btnSendToController.Caption = btnSendControllerCap
End Sub

Private Sub btnSendToController_Click()
    
    'ако контролера е ВИПА правим булевите променливи за проверка на данните му false
    If VipaActive = True Then
        WasAuto = False
        okAggr = False
        okCem = False
        okWat = False
        okHD = False
    Else
    End If
    
    On Error GoTo ErrorH
    
    'нулираме брояча на замесите
    CountMix = 0
    
    'стартираме таймер за започнала заявка
    DispPanel.TimerStartReq.Enabled = True
    
    'стартираме таймер за следене на резултати от замесите
    DispPanel.TimerRes.Enabled = True
    
    'подготвяме флаг за липсващи данни
    EmptyData = False

ErrorH:
    If Err.Number = 11 Then
' Messageshow
        Screen.MousePointer = vbDefault
        MsgBox MsgRestart, vbCritical
        Err.Clear
    End If
    
    'извикваме функцията за изпращаме към контролера
    Call SendToController
End Sub

Private Sub Form_Unload(Cancel As Integer)

    DispPanel.btnDispStart.Visible = True
    MousePointer = vbDefault
End Sub

