VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form prntForm3btn 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "����� 3"
   ClientHeight    =   13080
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11700
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   13080
   ScaleWidth      =   11700
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtExpNote 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   78
      TabStop         =   0   'False
      Top             =   1005
      Width           =   2175
   End
   Begin RichTextLib.RichTextBox Confirmity 
      Height          =   6015
      Left            =   360
      TabIndex        =   77
      Top             =   6840
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   10610
      _Version        =   393217
      ReadOnly        =   -1  'True
      TextRTF         =   $"prntForm3btn.frx":0000
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
      Height          =   375
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   3960
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
      Height          =   375
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   3600
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
      Height          =   375
      Left            =   7800
      Locked          =   -1  'True
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   3960
      Width           =   735
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
      Height          =   375
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   3600
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
      Height          =   375
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   3240
      Width           =   2175
   End
   Begin VB.TextBox txtOper 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   5400
      Width           =   2655
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
      Height          =   375
      Left            =   10680
      Locked          =   -1  'True
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   3960
      Width           =   735
   End
   Begin VB.TextBox txtEDM 
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
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   3960
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
      Height          =   375
      Left            =   9240
      Locked          =   -1  'True
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   3600
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
      Height          =   375
      Left            =   9240
      Locked          =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   3240
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
      Height          =   375
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   3240
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
      Height          =   375
      Left            =   10560
      Locked          =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2880
      Width           =   855
   End
   Begin VB.TextBox txtClass 
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
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2880
      Width           =   735
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
      Height          =   375
      Left            =   10320
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2520
      Width           =   615
   End
   Begin VB.TextBox txtRecType 
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
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2520
      Width           =   1935
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
      Height          =   375
      Left            =   10200
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2160
      Width           =   735
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
      Height          =   375
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2880
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
      Height          =   375
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2880
      Width           =   735
   End
   Begin VB.TextBox txtVol 
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
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2520
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
      Height          =   375
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2520
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
      Height          =   375
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2160
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
      Height          =   375
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2160
      Width           =   4695
   End
   Begin VB.TextBox txtObj 
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
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1800
      Width           =   4455
   End
   Begin VB.TextBox txtClnt 
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
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1800
      Width           =   4695
   End
   Begin VB.TextBox txtDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7320
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox txtOrd 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9240
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label lblConcP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lblConcP"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3840
      TabIndex        =   76
      Top             =   120
      Width           =   2595
   End
   Begin VB.Label lblFax 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lblFax"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6720
      TabIndex        =   75
      Top             =   360
      Width           =   2475
   End
   Begin VB.Label lblTel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lblTel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6720
      TabIndex        =   74
      Top             =   120
      Width           =   2445
   End
   Begin VB.Label lblTown 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lblTown"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   600
      TabIndex        =   73
      Top             =   360
      Width           =   2940
   End
   Begin VB.Label lblCompany 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lblCompany"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   600
      TabIndex        =   72
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label lblClassP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "������������:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3120
      TabIndex        =   71
      Top             =   3960
      Width           =   1320
   End
   Begin VB.Label lblCem2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2 - "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5640
      TabIndex        =   70
      Top             =   3240
      Width           =   285
   End
   Begin VB.Label lblCem1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1 - "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2520
      TabIndex        =   69
      Top             =   3240
      Width           =   285
   End
   Begin VB.Label lblOperSign 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "/..................."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4320
      TabIndex        =   68
      Top             =   5400
      Width           =   1200
   End
   Begin VB.Label lblMM 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "mm"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2520
      TabIndex        =   67
      Top             =   3960
      Width           =   330
   End
   Begin VB.Label lblChem3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "3 - "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   8880
      TabIndex        =   66
      Top             =   3600
      Width           =   285
   End
   Begin VB.Label lblChem2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2 - "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5640
      TabIndex        =   65
      Top             =   3600
      Width           =   285
   End
   Begin VB.Label lblChem1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1 - "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2520
      TabIndex        =   64
      Top             =   3600
      Width           =   285
   End
   Begin VB.Label lblCem3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "3 - "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   8880
      TabIndex        =   63
      Top             =   3240
      Width           =   285
   End
   Begin VB.Label lblM3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "m3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   11040
      TabIndex        =   62
      Top             =   2520
      Width           =   270
   End
   Begin VB.Label lblKG 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "kg"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   8040
      TabIndex        =   61
      Top             =   2520
      Width           =   210
   End
   Begin VB.Label lblM3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "m3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   5400
      TabIndex        =   60
      Top             =   2520
      Width           =   270
   End
   Begin VB.Label lblOrdVol 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���� �� ��������:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   8520
      TabIndex        =   59
      Top             =   2520
      Width           =   1710
   End
   Begin VB.Label lblRecType 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��� �������:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   360
      TabIndex        =   58
      Top             =   2520
      Width           =   1170
   End
   Begin VB.Label lblKM 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "km"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   11040
      TabIndex        =   57
      Top             =   2160
      Width           =   270
   End
   Begin VB.Label lblDist 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�����."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   9480
      TabIndex        =   56
      Top             =   2160
      Width           =   570
   End
   Begin VB.Label lblAdd 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   $"prntForm3btn.frx":0084
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   360
      TabIndex        =   55
      Top             =   5880
      Width           =   11070
   End
   Begin VB.Label lblAddSign 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "������������: ........................................"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7680
      TabIndex        =   54
      Top             =   6240
      Width           =   3780
   End
   Begin VB.Label lblAddType 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���: ..................................."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4200
      TabIndex        =   53
      Top             =   6240
      Width           =   2580
   End
   Begin VB.Label lblAddPlace 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�����: ................................................"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   360
      TabIndex        =   52
      Top             =   6240
      Width           =   3555
   End
   Begin VB.Label lblSign 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�����: ....................................."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   8520
      TabIndex        =   51
      Top             =   5400
      Width           =   2895
   End
   Begin VB.Label lblDrvSign 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�����: ........................."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6120
      TabIndex        =   50
      Top             =   5400
      Width           =   2190
   End
   Begin VB.Label lblOper 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��������:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   360
      TabIndex        =   49
      Top             =   5400
      Width           =   1020
   End
   Begin VB.Label lblPlaceEnd 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���� �� ��������: ........................................"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6120
      TabIndex        =   48
      Top             =   4920
      Width           =   4095
   End
   Begin VB.Label lblTimeArr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��� �� ����������: ......................................"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6120
      TabIndex        =   47
      Top             =   4560
      Width           =   4065
   End
   Begin VB.Label lblChem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��������� �������:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   46
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label lblEDM 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "D max �� ���:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   45
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label lblCem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���� ������:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   360
      TabIndex        =   44
      Top             =   3240
      Width           =   1185
   End
   Begin VB.Label lblClassH 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���� �� �-��� �� �������:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   8040
      TabIndex        =   43
      Top             =   2880
      Width           =   2385
   End
   Begin VB.Label lblClassV 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���� �� ����.:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5520
      TabIndex        =   42
      Top             =   2880
      Width           =   1320
   End
   Begin VB.Label lblClassK 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���� �� �������.:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2880
      TabIndex        =   41
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label lblClass 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���� �� �����:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   360
      TabIndex        =   40
      Top             =   2880
      Width           =   1305
   End
   Begin VB.Label lblPlaceStart 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "������ �� ����������: ...................................."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   360
      TabIndex        =   39
      Top             =   4920
      Width           =   4260
   End
   Begin VB.Label lblPlaceMix 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�������������: .............................................................."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   360
      TabIndex        =   38
      Top             =   4560
      Width           =   5205
   End
   Begin VB.Label lblW 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�����:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6240
      TabIndex        =   37
      Top             =   2520
      Width           =   555
   End
   Begin VB.Label lblVol 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3840
      TabIndex        =   36
      Top             =   2520
      Width           =   555
   End
   Begin VB.Label lblExpTime 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��� �� ����������:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   8760
      TabIndex        =   35
      Top             =   3960
      Width           =   1785
   End
   Begin VB.Label lblMixTime 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "������ �� ���������:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5640
      TabIndex        =   34
      Top             =   3960
      Width           =   2010
   End
   Begin VB.Label lblDrvNo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6240
      TabIndex        =   33
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label lblDrv 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�����:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   360
      TabIndex        =   32
      Top             =   2160
      Width           =   630
   End
   Begin VB.Label lblObj 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�����:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6240
      TabIndex        =   31
      Top             =   1800
      Width           =   585
   End
   Begin VB.Label lblClnt 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "������:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   360
      TabIndex        =   30
      Top             =   1800
      Width           =   675
   End
   Begin VB.Label lblOrd 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�� ������ No. / ����"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   9240
      TabIndex        =   29
      Top             =   840
      Width           =   2115
   End
   Begin VB.Label lblDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   7080
      TabIndex        =   28
      Top             =   1080
      Width           =   105
   End
   Begin VB.Label lblExpNote 
      BackStyle       =   0  'Transparent
      Caption         =   "������������� ������� No."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   27
      Top             =   1080
      Width           =   4215
   End
   Begin VB.Label Label35 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "�������� ���-����� v1.0b/2013"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   9360
      TabIndex        =   26
      Top             =   120
      Width           =   2145
   End
End
Attribute VB_Name = "prntForm3btn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'��������� �� ����������� �� ����� 3 �� ���������
    Dim PrevSet As Boolean
    Dim strSubKey As String
    strSubKey = Trim(PlaceProgSet3)
    PrevSet = CheckRegistryKey(HKEY_CURRENT_USER, strSubKey)
    
    If PrevSet = True Then
        rDist = GetSetting(PlaceProgSettings, PlaceForm3, "Dist", ErrRes)
        rRecType = GetSetting(PlaceProgSettings, PlaceForm3, "RecType", ErrRes)
        rVol = GetSetting(PlaceProgSettings, PlaceForm3, "Vol", ErrRes)
        rW = GetSetting(PlaceProgSettings, PlaceForm3, "W", ErrRes)
        rOrdVol = GetSetting(PlaceProgSettings, PlaceForm3, "OrdVol", ErrRes)
        rClass = GetSetting(PlaceProgSettings, PlaceForm3, "Class", ErrRes)
        rClassK = GetSetting(PlaceProgSettings, PlaceForm3, "ClassK", ErrRes)
        rClassV = GetSetting(PlaceProgSettings, PlaceForm3, "ClassV", ErrRes)
        rClassH = GetSetting(PlaceProgSettings, PlaceForm3, "ClassH", ErrRes)
        rClassP = GetSetting(PlaceProgSettings, PlaceForm3, "ClassP", ErrRes)
        rCem1 = GetSetting(PlaceProgSettings, PlaceForm3, "Cem1", ErrRes)
        rCem2 = GetSetting(PlaceProgSettings, PlaceForm3, "Cem2", ErrRes)
        rCem3 = GetSetting(PlaceProgSettings, PlaceForm3, "Cem3", ErrRes)
        rChem1 = GetSetting(PlaceProgSettings, PlaceForm3, "Chem1", ErrRes)
        rChem2 = GetSetting(PlaceProgSettings, PlaceForm3, "Chem2", ErrRes)
        rChem3 = GetSetting(PlaceProgSettings, PlaceForm3, "Chem3", ErrRes)
        rEDM = GetSetting(PlaceProgSettings, PlaceForm3, "EDM", ErrRes)
        rMixTime = GetSetting(PlaceProgSettings, PlaceForm3, "MixTime", ErrRes)
        rExpTime = GetSetting(PlaceProgSettings, PlaceForm3, "ExpTime", ErrRes)
        rRealVol = GetSetting(PlaceProgSettings, PlaceForm3, "RealVol", ErrRes)
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
    
    
    Dim Comp As String
    Dim Town As String
    Dim ConcP As String
    Dim Tel As String
    Dim Fax As String
    Dim intEmpFileNbr1 As Integer
    intEmpFileNbr1 = FreeFile
    
    If Dir(InfoFile) <> "" Then
        Open InfoFile For Input As intEmpFileNbr1
        Input #intEmpFileNbr1, Comp, Town, ConcP, Tel, Fax
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
            Me.lblTel.Caption = uniTel & ": " & Tel
        Else
            Me.lblTel.Caption = ""
        End If
        If Len(Fax) Then
            Me.lblFax.Caption = uniFax & ": " & Fax
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
    
    If Dir(ConfirmityFile) <> "" Then
        Me.Confirmity.LoadFile ConfirmityFile, 1
    End If
End Sub


